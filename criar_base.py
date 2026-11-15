"""Cria uma planilha-exemplo."""
import pandas as pd

dados = [
    ("Embrapa", "https://www.embrapa.br", "Brasil"),
    ("Syngenta Brasil", "https://www.syngenta.com.br", "Brasil"),
    ("BASF Agro", "https://www.agro.basf.com.br", "Brasil"),
    ("Bayer Crop Science", "https://www.cropscience.bayer.com.br", "Brasil"),
    ("Corteva Agriscience", "https://www.corteva.com.br", "Brasil"),
    ("John Deere Brasil", "https://www.deere.com.br", "Brasil"),
    ("Yara Brasil", "https://www.yara.com.br", "Brasil"),
    ("Nutrien", "https://www.nutrien.com", "Canadá"),
    ("Cargill", "https://www.cargill.com.br", "Brasil"),
    ("Bunge", "https://www.bunge.com.br", "Brasil"),
    ("Bioceres Crop Solutions", "https://www.biocerescrops.com", "Argentina"),
    ("Nufarm", "https://nufarm.com", "Austrália"),
    ("Agrosul Insumos Ltda", "http://www.agrosulinsumos-exemplo.com.br", "Brasil"),      # domínio inexistente
    ("Fazenda Boa Safra S.A.", "http://www.fazendaboasafra-exemplo.com.br", "Brasil"),   # domínio inexistente
    ("Cooperativa Vale Verde", "https://httpbin.org/status/404", "Brasil"),              # simula erro 404
    ("AgroTech Solutions", "https://httpbin.org/status/500", "Paraguai"),                # simula erro 500
    ("Sementes Horizonte", "https://httpbin.org/delay/15", "Uruguai"),                   # simula timeout
    ("Grupo Campo Forte", "https://example.com", "Brasil"),                              # online, mas título diferente
    ("Agro Nova Mudas", "", "Brasil"),                                                   # sem website
    ("Frigorífico Pampa", "www.pampa-exemplo.com.br", "Brasil"),                         # sem http:// no cadastro
]

df = pd.DataFrame(dados, columns=["Nome da Empresa", "Website", "País"])
df["Status_Verificação"] = "Pendente"
df.to_excel("base_exemplo.xlsx", index=False)
print(f"Base criada com {len(df)} empresas.")
