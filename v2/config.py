# -- CONFIGURAÇÕES DO RELATÓRIO --

MES = "jan"
ANO = "26"
NOME_RELATORIO = "BEG001"

# -- PASTA DE DOWNLOAD (CAMINHO FIXO) --
PASTA_DOWNLOAD = r"C:\Users\malik.mourad\Downloads"

# -- VALIDAÇÃO DE URL (SEGURANÇA) --
# Lista de URLs/domínios válidos da CCEE
URL_ESPERADA_CCEE = [
    "ccee.org.br",
    "bi.ccee.org.br",
    "analytics",
    "saw.dll",  # ← vírgula aqui!
    "operacao.ccee.org.br",
    "https://operacao.ccee.org.br/ui/dri/dashboard"
]

# -- VALIDAÇÃO DE CONTEÚDO --
# XPaths de elementos que DEVEM existir em um relatório válido
# Ajuste conforme a estrutura do seu relatório específico
ELEMENTOS_VALIDACAO = [
    "//table[contains(@class, 'TableView')]",  # Tabela de dados
    "//div[contains(@class, 'ViewContainer')]",  # Container de visualização
    "//*[contains(@class, 'DashboardContent')]",  # Conteúdo do dashboard
    "//td[contains(@class, 'TableHeaderCell')]",  # Cabeçalho de tabela
    "//*[contains(@id, 'saw_')]",  # IDs típicos do Oracle BI
    "//div[contains(@class, 'ResultsContainer')]"  # Container de resultados
]

# -- LISTA DE EMPRESAS --
empresas_alvo = [
    "ADESI CL 514",
    "AMERICAS CL 514",
    "AOI YAMA CL 514",
    "APOLO TUBULARS",
    "APUCARANINHA",
    "ARBHORES COMPENSADOS CL 514",
    "AURORA MATRIZ",
    "BAIA VERDE",
    "CAVERNOSO I",
    "CAVERNOSO II",
    "C VALE COOP AGROINDUSTRIAL",
    "C VALE ENERGIA",
    "CERMISSOES",
    "CERT",
    "CERTAJA",
    "CERTHIL DISTRIBUICAO",
    "CGH DONA MARIA PIANA",
    "CGH FAXINAL DOS GUEDES",
    "CGH MARUMBI ESP",
    "CHAMINE",
    "CHOPIM 1 ESP",
    "CIFA CL 514",
    "CONTESTADO",
    "CONTINENTAL COM",
    "COOPERALFA",
    "COOPERLUZ",
    "COOPERLUZ DIST",
    "CORONEL ARAUJO",
    "CRERAL DIST",
    "CVALE",
    "DETOFOL",
    "ELECTRA ENERGIA DIGITAL",
    "ELECTRA ENERGY",
    "EMBAUBA",
    "GUARICANA I5",
    "ISOELECTRIC L",
    "ITAGUACU",
    "MARCEGAGLIA",
    "MELISSA ESP PI",
    "MET CARTEC CL 514",
    "PALMAS",
    "PASSO FERRAZ",
    "PCH BURITI",
    "PCH CAMBARA",
    "PCH FERRADURA",
    "PCH PESQUEIRO",
    "PCH PONTE BRANCA",
    "PCH QUEBRA DENTES",
    "PCH RIO INDIOS",
    "PCH SALTO DO GUASSUPI",
    "PCH TAMBAU",
    "PEQUI I5",
    "PITANGUI ESP PI",
    "PLASTBEL",
    "PLASTICOS PR",
    "RINCAO DOS ALBINOS",
    "RINCAO SAO MIGUEL",
    "SALTO DO TIMBO",
    "SALTO DO VAU I5",
    "SAO JORGE ESPECIAL",
    "SJHOTEIS",
    "SUCUPIRA",
    "THERMOPLAST CL 514",
    "TIMBER CL 514",
    "USINAS DO PRATA",
    "UTE GOIANESIA",
    "VIDROLAR CL 514",
    "VITORINO",
    "ZAELI CL 514",
]
"""
empresas_alvo = [ 
    "ADESI CL 514",
    "AMERICAS CL 514",
    "AOI YAMA CL 514",
    "APOLO TUBULARS",
    "ARBHORES COMPENSADOS CL 514",
    "AURORA FASGO CL",
    "AURORA MATRIZ",
    "AURORA XAXIM",
    "BAIA VERDE",
    "C VALE COOP AGROINDUSTRIAL",
    "CERMISSOES",
    "CERTAJA",
    "CERTHIL DISTRIBUICAO",
    "CIFA CL 514",
    "COOPERALFA",
    "COOPERLUZ DIST",
    "CRERAL DIST",
    "ELECTRA ENERGIA DIGITAL",
    "ISOELECTRIC L",
    "MARCEGAGLIA",
    "MET CARTEC CL 514",
    "PLASTBEL",
    "PLASTICOS PR",
    "SJHOTEIS",
    "THERMOPLAST CL 514",
    "TIMBER CL 514",
    "UTE GOIANESIA",
    "VIDROLAR CL 514",
    "ZAELI CL 514",
]
"""