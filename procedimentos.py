import pandas as pd


def carregar(procedimento):
    arquivo_excel = "db.xlsx"
    tabela = pd.read_excel(arquivo_excel)
    coluna = [x for x in tabela[procedimento].to_list() if pd.notna(x) and x not in [None, '']] 
    return coluna

def pacote_otorrino():
    return carregar("PACOTE PRÉ-OPERATÓRIO PEDIÁTRICO OTORRINO")

def pacote_geral():
    return carregar("PACOTE PRÉ-OPERATÓRIO PEDIÁTRICO CIRURGIA GERAL")

def pacote_oftalmo():
    return carregar("PACOTE PRÉ-OPERATÓRIO PEDIÁTRICO OFTALMOLOGISTA")

def pacote_adeno():
    return carregar("ADENOIDECTOMIA PEDIÁTRICO")

def pacote_amig():
    return carregar("AMIGDALECTOMIA - PEDIATRICO")

def pacote_amig_adeno():
    return carregar("AMIGDALECTOMIA COM ADENOIDECTOMIA - PEDIATRICO")

def pacote_nasal():
    return carregar("TRATAMENTO CIRÚRGICO DE PERFURAÇÃO DO SEPTO NASAL - PEDIATRICO")

def pacote_estrabismo():
    return carregar("CORREÇÃO CIRÚRGICA DE ESTRABISMO (ACIMA DE 2 MUSCULOS) - PEDIATRICO")

def pacote_inguinal():
    return carregar("HERNIOPLASTIA INGUINAL (BILATERAL) - PEDIATRICO")

def pacote_umbilical():
    return carregar("HERNIOPLASTIA UMBILICAL - PEDIATRICO")

def pacote_orqui():
    return carregar("ORQUIDOPEXIA BILATERAL - PEDIATRICO")

def pacote_hidrocele():
    return carregar("TRATAMENTO CIRÚRGICO DE HIDROCELE - PEDIATRICO")

def pacote_hispospadia():
    return carregar("CORRECAO DE HIPOSPADIA (1º TEMPO) - PEDIATRICO")

def pacote_plastica():
    return carregar("PLASTICA TOTAL DO PENIS - PEDIATRICO")

def pacote_postec():
    return carregar("POSTECTOMIA - PEDIATRICO")
