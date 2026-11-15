# Validador de Perfis de Empresas

Automação em Python que verifica se os websites de uma base de empresas ainda estão no ar e se o nome da empresa ainda aparece neles. O resultado é uma planilha Excel com o status de cada cadastro e um gráfico-resumo.
---

## Sumário

- [O problema](#o-problema)
- [O que o projeto faz](#o-que-o-projeto-faz)
- [Estrutura do projeto](#estrutura-do-projeto)
- [Instalação](#instalação)
- [Como usar](#como-usar)
- [Regras de classificação](#regras-de-classificação)
- [Saídas geradas](#saídas-geradas)
- [Resultado da simulação](#resultado-da-simulação)
- [Limitações conhecidas](#limitações-conhecidas)
- [Tecnologias](#tecnologias)

---

## O problema

A qualidade dos dados de contato, faixa de receita e categoria é o principal ativo de uma base de empresas. Com o tempo, sites saem do ar, domínios expiram e empresas mudam de nome. Verificar tudo isso à mão, linha por linha, leva semanas e não escala.

## O que o projeto faz

1. Lê uma planilha Excel com as empresas (nome, website, país).
2. Acessa o website de cada empresa por HTTP, de forma paralela.
3. Classifica cada cadastro em um status (veja [as regras](#regras-de-classificação)).
4. Extrai o `<title>` da página e confere se o nome da empresa ainda bate.
5. Gera uma nova planilha Excel com os resultados, uma aba de resumo com gráfico e um gráfico em PNG.

## Estrutura do projeto

```
.
├── criar_base.py                          # gera a base de exemplo (20 empresas)
├── verificar_sites.py                     # script principal de verificação
├── base_exemplo.xlsx                      # entrada: base simulada
├── resultado_verificacao.xlsx             # saída: resultados + resumo + gráfico
├── resultado_verificacao_grafico.png      # saída: gráfico (matplotlib)
└── README.md
```

## Instalação

Requer **Python 3.9 ou superior**.

```bash
pip install pandas requests openpyxl matplotlib
```

## Como usar

**1. Criar a base de exemplo** (opcional, se você já tiver sua própria planilha):

```bash
python criar_base.py
```

**2. Rodar a verificação real** (precisa de internet):

```bash
python verificar_sites.py
```

**3. Rodar em modo demonstração** (sem internet, com respostas simuladas):

```bash
python verificar_sites.py --demo
```

### Opções da linha de comando

| Opção | Descrição | Padrão |
|---|---|---|
| `-i`, `--entrada` | Planilha de entrada | `base_exemplo.xlsx` |
| `-o`, `--saida` | Planilha de saída | `resultado_verificacao.xlsx` |
| `--demo` | Usa respostas HTTP simuladas | desligado |
| `--workers` | Número de verificações simultâneas | `10` |

### Formato esperado da planilha de entrada

| Nome da Empresa |        Website         |  País  | Status_Verificação |
|-----------------|------------------------|--------|--------------------|
|     Embrapa     | https://www.embrapa.br | Brasil |      Pendente      |

As colunas `Nome da Empresa` e `Website` são obrigatórias. URLs sem `http://` ou `https://` são completadas automaticamente.

## Regras de classificação

| Situação encontrada | Status gerado |
|---|---|
| Erro 4xx/5xx (ex.: 404, 500), timeout, falha de DNS ou erro de SSL | `Site Inativo / Verificar` |
| Código 401, 403 ou 429 (servidor respondeu, mas bloqueia robôs) | `Ativo - Acesso restrito / Conferir manualmente` |
| Site online e o `<title>` contém o nome da empresa | `Ativo - Nome confere` |
| Site online, mas o `<title>` não contém o nome da empresa | `Ativo - Nome divergente / Revisar` |
| Campo Website vazio | `Sem Website / Verificar` |

**Como o nome é comparado:** o script remove acentos, pontuação e termos genéricos (como "Ltda", "S.A.", "Brasil") e verifica se pelo menos **50% das palavras-chave** do nome aparecem no título da página ou no domínio.

Sites com acesso restrito não são marcados como inativos de propósito: o servidor está no ar, só não aceita acesso automatizado. Marcá-los como inativos geraria falsos alarmes.

## Saídas geradas

**Planilha Excel** (`resultado_verificacao.xlsx`):

- **Aba `Resumo`:** total de empresas, ativas vs. inativas, detalhamento por status e gráfico de colunas nativo do Excel. Os números usam fórmulas (`COUNTIF`) e se atualizam se os status forem editados.
- **Aba `Resultados`:** a base original mais as colunas `Código_HTTP`, `Título_da_Página`, `Observação` e `Data_Verificação`. Os status são coloridos, com filtro e cabeçalho fixo.

**Gráfico PNG** (`..._grafico.png`): barras de empresas ativas vs. inativas, feito com matplotlib.

## Resultado da simulação

> **Atenção:** os números abaixo vêm do **modo `--demo`**, com respostas HTTP simuladas. Não representam o estado real desses sites. Para dados reais, rode o script sem `--demo`.

Base de 20 empresas:

| Categoria | Quantidade | % |
|---|---|---|
| Ativas | 13 | 65% |
| Inativas / a verificar | 7 | 35% |

Detalhamento: 11 com nome confere, 1 com nome divergente, 1 com acesso restrito, 6 inativos e 1 sem website.

## Limitações conhecidas

- **Sites com conteúdo carregado por JavaScript** podem ter `<title>` genérico ("Home", "Welcome") e cair em "nome divergente". Por isso esse status pede **revisão**, não exclusão.
- **Bloqueio anti-robô:** alguns sites recusam acessos automatizados mesmo estando no ar.
- **Falha momentânea:** um site pode estar fora do ar só na hora da verificação, e hoje há uma única tentativa por site.
- **Comparação de nome simples:** não reconhece siglas, nomes fantasia ou marcas diferentes da razão social.

## Tecnologias

- **Python**
- **pandas:** leitura e manipulação da planilha
- **requests:** requisições HTTP
- **openpyxl:** formatação do Excel, fórmulas e gráfico nativo
- **matplotlib:** gráfico em PNG
- **concurrent.futures:** verificação em paralelo
