"""
Validador de Perfis de Empresas
-------------------------------------------
Lê a planilha-base, acessa o website de cada empresa, classifica o status
e gera uma nova planilha Excel (com aba de resumo + gráfico) e um PNG do gráfico.
"""

import argparse
import html
import re
import unicodedata
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime

import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.chart import BarChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

TIMEOUT = 10
HEADERS = {"User-Agent": "Mozilla/5.0 (compatible; validator/1.0)"}

ST_INATIVO = "Site Inativo / Verificar"
ST_OK = "Ativo - Nome confere"
ST_DIVERGENTE = "Ativo - Nome divergente / Revisar"
ST_RESTRITO = "Ativo - Acesso restrito / Conferir manualmente"
ST_SEM_SITE = "Sem Website / Verificar"

STOPWORDS = {"ltda", "sa", "s", "a", "me", "epp", "de", "da", "do", "das", "dos", "e",
             "brasil", "brazil", "grupo", "cia", "company", "inc", "ltd", "the"}

# utilidades
def normalizar(txt: str) -> str:
    txt = unicodedata.normalize("NFKD", str(txt)).encode("ascii", "ignore").decode()
    return re.sub(r"[^a-z0-9 ]", " ", txt.lower())


def tokens_nome(nome: str) -> list[str]:
    return [t for t in normalizar(nome).split() if t not in STOPWORDS and len(t) >= 3]


def nome_confere(nome: str, titulo: str, url: str) -> bool:
    """True se >= 50% das palavras-chave do nome aparecem no <title> ou no domínio."""
    toks = tokens_nome(nome)
    if not toks:
        return False
    alvo = normalizar(titulo) + " " + normalizar(url).replace(" ", "")
    achou = sum(1 for t in toks if t in alvo)
    return achou / len(toks) >= 0.5


def extrair_titulo(texto: str) -> str:
    m = re.search(r"<title[^>]*>(.*?)</title>", texto[:300_000], re.I | re.S)
    return re.sub(r"\s+", " ", html.unescape(m.group(1))).strip() if m else ""


def limpar_url(url) -> str:
    url = "" if pd.isna(url) else str(url).strip()
    if url and not url.startswith(("http://", "https://")):
        url = "https://" + url
    return url


# consulta HTTP
def consultar_real(url: str):
    """Retorna (codigo_http, titulo, erro)."""
    try:
        r = requests.get(url, headers=HEADERS, timeout=TIMEOUT, allow_redirects=True)
        titulo = extrair_titulo(r.text) if r.status_code < 400 else ""
        return r.status_code, titulo, None
    except requests.exceptions.Timeout:
        return None, "", "Timeout"
    except requests.exceptions.SSLError:
        return None, "", "Erro de certificado SSL"
    except requests.exceptions.ConnectionError:
        return None, "", "Falha de conexão / DNS"
    except requests.exceptions.RequestException as e:
        return None, "", f"Erro: {type(e).__name__}"

# classificação
def verificar_linha(nome: str, url_original):
    url = limpar_url(url_original)
    if not url:
        return dict(status=ST_SEM_SITE, codigo=None, titulo="", obs="Website não informado")

    codigo, titulo, erro = consultar_real(url)

    if erro:
        return dict(status=ST_INATIVO, codigo=None, titulo="", obs=erro)
    if codigo in (401, 403, 429):
        return dict(status=ST_RESTRITO, codigo=codigo, titulo="",
                    obs="Servidor respondeu, mas bloqueia acesso automatizado")
    if codigo >= 400:
        return dict(status=ST_INATIVO, codigo=codigo, titulo="", obs=f"HTTP {codigo}")

    if nome_confere(nome, titulo, url):
        return dict(status=ST_OK, codigo=codigo, titulo=titulo, obs="")
    return dict(status=ST_DIVERGENTE, codigo=codigo, titulo=titulo,
                obs="Título da página não contém o nome da empresa")


# saídas
def gerar_grafico_png(df: pd.DataFrame, caminho: str):
    ativas = df["Status_Verificação"].str.startswith("Ativo").sum()
    inativas = len(df) - ativas
    fig, ax = plt.subplots(figsize=(6, 4.2))
    barras = ax.bar(["Ativas", "Inativas / A verificar"], [ativas, inativas],
                    color=["#2E7D32", "#C62828"])
    ax.bar_label(barras, fontsize=12, fontweight="bold")
    ax.set_ylabel("Nº de empresas")
    ax.set_title("Empresas com website ativo vs. inativo", fontsize=11)
    ax.spines[["top", "right"]].set_visible(False)
    fig.tight_layout()
    fig.savefig(caminho, dpi=150)
    plt.close(fig)


def formatar_excel(caminho: str):
    wb = load_workbook(caminho)
    ws = wb["Resultados"]
    fonte = Font(name="Arial", size=10)
    cab_fill = PatternFill("solid", start_color="1B5E20")
    cores = {ST_OK: "C8E6C9", ST_DIVERGENTE: "FFF9C4",
             ST_RESTRITO: "FFE0B2", ST_INATIVO: "FFCDD2", ST_SEM_SITE: "E0E0E0"}

    for cel in ws[1]:
        cel.font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        cel.fill = cab_fill
        cel.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    col_status = [c.value for c in ws[1]].index("Status_Verificação") + 1
    for row in ws.iter_rows(min_row=2):
        for cel in row:
            cel.font = fonte
        cor = cores.get(row[col_status - 1].value)
        if cor:
            row[col_status - 1].fill = PatternFill("solid", start_color=cor)
    larguras = [28, 42, 12, 44, 12, 48, 40, 18]
    for i, w in enumerate(larguras, 1):
        ws.column_dimensions[get_column_letter(i)].width = w
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions

    # aba Resumo (fórmulas + gráfico nativo do Excel)
    rs = wb.create_sheet("Resumo")
    n = ws.max_row
    rng = f"Resultados!$D$2:$D${n}"
    rs["A1"] = "Resumo da verificação de websites"
    rs["A1"].font = Font(name="Arial", size=13, bold=True)
    rs["A2"] = "Gerado em: " + datetime.now().strftime("%d/%m/%Y %H:%M")
    rs["A2"].font = Font(name="Arial", size=9, italic=True, color="555555")

    # Bloco 1: ativas vs. inativas (alimenta o gráfico)
    rs["A4"], rs["B4"] = "Total de empresas", f"=COUNTA(Resultados!$A$2:$A${n})"
    rs["A6"], rs["B6"], rs["C6"] = "Categoria", "Nº de empresas", "% do total"
    rs["A7"], rs["B7"] = "Ativas", f'=COUNTIF({rng},"Ativo*")'
    rs["A8"], rs["B8"] = "Inativas / A verificar", "=B4-B7"
    
    
    # Bloco 2: detalhamento por status
    rs["A10"], rs["B10"], rs["C10"] = "Detalhamento por status", "Nº de empresas", "% do total"

    detalhes = [ST_OK, ST_DIVERGENTE, ST_RESTRITO, ST_INATIVO, ST_SEM_SITE]

    for i, st in enumerate(detalhes, start=11):
        rs[f"A{i}"], rs[f"B{i}"] = st, f'=COUNTIF({rng},"{st}")'

    for r in [7, 8] + list(range(11, 11 + len(detalhes))):
        rs[f"C{r}"] = f"=IF($B$4=0,0,B{r}/$B$4)"
        rs[f"C{r}"].number_format = "0.0%"

    for row in rs.iter_rows(min_row=4, max_row=15):
        for cel in row:
            cel.font = Font(name="Arial", size=10)

    rs["A4"].font = rs["B4"].font = Font(name="Arial", size=10, bold=True)

    for cel in ("A6", "B6", "C6", "A10", "B10", "C10"):
        rs[cel].font = Font(name="Arial", size=10, bold=True, color="FFFFFF")
        rs[cel].fill = cab_fill
        
    rs.column_dimensions["A"].width = 48
    rs.column_dimensions["B"].width = 16
    rs.column_dimensions["C"].width = 12

    ch = BarChart()
    ch.type = "col"
    ch.title = "Empresas ativas vs. inativas"
    ch.y_axis.title = "Nº de empresas"
    ch.add_data(Reference(rs, min_col=2, min_row=6, max_row=8), titles_from_data=True)
    ch.set_categories(Reference(rs, min_col=1, min_row=7, max_row=8))
    ch.legend = None
    ch.dataLabels = DataLabelList()
    ch.dataLabels.showVal = True
    ch.roundedCorners = False
    ch.x_axis.delete = False
    ch.y_axis.delete = False
    ch.height, ch.width = 8, 12
    rs.add_chart(ch, "E4")
    wb.move_sheet("Resumo", offset=-1)
    wb.save(caminho)


# principal
def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-i", "--entrada", default="base_exemplo.xlsx")
    ap.add_argument("-o", "--saida", default="resultado_verificacao.xlsx")
    ap.add_argument("--workers", type=int, default=10)
    args = ap.parse_args()

    df = pd.read_excel(args.entrada)

    with ThreadPoolExecutor(max_workers=args.workers) as ex:
        res = list(ex.map(lambda r: verificar_linha(r["Nome da Empresa"], r["Website"]),
                          df.to_dict("records")))

    df["Status_Verificação"] = [r["status"] for r in res]
    df["Código_HTTP"] = [r["codigo"] for r in res]
    df["Título_da_Página"] = [r["titulo"] for r in res]
    df["Observação"] = [r["obs"] for r in res]
    df["Data_Verificação"] = datetime.now().strftime("%d/%m/%Y")

    df.to_excel(args.saida, index=False, sheet_name="Resultados")
    formatar_excel(args.saida)
    png = args.saida.replace(".xlsx", "_grafico.png")
    gerar_grafico_png(df, png)

    print(df["Status_Verificação"].value_counts().to_string())
    print(f"\nArquivos gerados:\n - {args.saida}\n - {png}")


if __name__ == "__main__":
    main()
