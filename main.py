from Prestador.neotin.neotin import analisar_neotin
from Prestador.neomater.neomater import analisar_neomater
from Prestador.prontobaby.prontobaby import analisar_prontobaby
from separarRelatorio.main import processar_todos_arquivos_simplificado
# from segvision.segvision import analisar_segvision

if __name__ == "__main__":
    processar_todos_arquivos_simplificado()
    
    analisar_neotin()
    analisar_neomater()
    analisar_prontobaby()
    # analisar_segvision()