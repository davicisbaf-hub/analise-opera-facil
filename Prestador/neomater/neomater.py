import pandas as pd
from procedimentos import (
    pacote_otorrino, pacote_geral, pacote_oftalmo, pacote_hispospadia,
    pacote_inguinal, pacote_hidrocele, pacote_adeno, pacote_amig,
    pacote_amig_adeno, pacote_estrabismo, pacote_nasal, pacote_orqui,
    pacote_plastica, pacote_postec, pacote_umbilical
)

def analisar_neomater():
    arquivo = "relatorios_simplificados/separarNeomater_SIMPLIFICADO.xlsx"
    municipios = ["RJ - Belford Roxo", "RJ - Duque de Caxias", "RJ - Itaguaí", "RJ - Japeri", "RJ - Magé", "RJ - Mesquita", "RJ - Nilópolis", "RJ - Nova Iguaçu", "RJ - Paracambi", "RJ - Queimados", "RJ - Seropédica", "RJ - São João de Meriti"]
    
    # Dicionário para acumular os resultados de todos os municípios
    resultados_por_municipio = {}
    
    # Lista para acumular DataFrames de procedimentos não listados
    nao_listados_dfs = []

    for municipio in municipios:
        try:
            # Inicializar todas as listas de procedimentos
            otorrino = pacote_otorrino()
            geral = pacote_geral()
            oftalmo = pacote_oftalmo()
            hispospadia = pacote_hispospadia()
            inguinal = pacote_inguinal()
            hidrocele = pacote_hidrocele()
            adeno = pacote_adeno()
            amig = pacote_amig()
            amig_adeno = pacote_amig_adeno()
            estrabismo = pacote_estrabismo()
            nasal = pacote_nasal()
            orqui = pacote_orqui()
            plastica = pacote_plastica()
            postec = pacote_postec()
            umbilical = pacote_umbilical()

            tabela = pd.read_excel(arquivo)

            coluna_procedimento = municipio
            coluna_quantidade = f"Quantidade {municipio}"

            if coluna_procedimento not in tabela.columns:
                print(f"Coluna '{coluna_procedimento}' não encontrada!")
                continue
            
            if coluna_quantidade not in tabela.columns:
                print(f"Coluna '{coluna_quantidade}' não encontrada!")
                continue
            
            resultados_municipio = {municipio: {} for municipio in municipios}  # ← Todos municipios já incluídos   
            
            def process_group(procedimentos):
                total = 0
                for procedimento in procedimentos:
                    mask = tabela[coluna_procedimento].astype(str) == str(procedimento)
                    quantidade = tabela.loc[mask, coluna_quantidade].sum()
                    
                    try:
                        quantidade_num = float(quantidade) if not pd.isna(quantidade) else 0
                    except (ValueError, TypeError):
                        quantidade_num = 0

                    if quantidade_num > 0:
                        resultados_municipio[str(procedimento)] = quantidade_num
                        total += quantidade_num
                return total
            
            # Grupos de procedimentos
            grupos = {
                "PACOTE PRÉ-OPERATÓRIO PEDIÁTRICO OTORRINO": otorrino,
                "PACOTE PRÉ-OPERATÓRIO PEDIÁTRICO CIRURGIA GERAL": geral,
                "PACOTE PRÉ-OPERATÓRIO PEDIÁTRICO OFTALMOLOGISTA": oftalmo,
                "ADENOIDECTOMIA PEDIÁTRICO": adeno,
                "AMIGDALECTOMIA - PEDIATRICO": amig,
                "AMIGDALECTOMIA COM ADENOIDECTOMIA - PEDIATRICO": amig_adeno,
                "TRATAMENTO CIRÚRGICO DE PERFURAÇÃO DO SEPTO NASAL - PEDIATRICO": nasal,
                "CORREÇÃO CIRÚRGICA DE ESTRABISMO (ACIMA DE 2 MUSCULOS) - PEDIATRICO": estrabismo,
                "HERNIOPLASTIA INGUINAL (BILATERAL) - PEDIATRICO": inguinal,
                "HERNIOPLASTIA UMBILICAL - PEDIATRICO": umbilical,
                "ORQUIDOPEXIA BILATERAL - PEDIATRICO": orqui,
                "TRATAMENTO CIRÚRGICO DE HIDROCELE - PEDIATRICO": hidrocele,
                "CORRECAO DE HIPOSPADIA (1º TEMPO) - PEDIATRICO": hispospadia,
                "PLASTICA TOTAL DO PENIS - PEDIATRICO": plastica,
                "POSTECTOMIA - PEDIATRICO": postec,
            }
            
            totais = {}
            for nome, lista in grupos.items():
                totais[nome] = process_group(lista or [])
        
            resultados_por_municipio[municipio] = totais
            
            print(f"=== RESULTADOS {municipio} ===")
            for nome, total in totais.items():
                print(f"{nome}: {total}")
            soma_total = sum(totais.values())
            print(f"Soma Total: {soma_total}")
            print("\n")
            
            # Procedimentos não listados
            todos_procedimentos_conhecidos = []
            for lista in grupos.values():
                todos_procedimentos_conhecidos.extend(lista or [])
            
            todos_procedimentos_conhecidos = list(set([str(p) for p in todos_procedimentos_conhecidos if p and str(p).strip()]))
            
            procedimentos_na_coluna = tabela[coluna_procedimento].dropna().unique()
            procedimentos_na_coluna = [str(p) for p in procedimentos_na_coluna if p and str(p).strip()]
            
            procedimentos_nao_mapeados = {}
            for procedimento in procedimentos_na_coluna:
                if procedimento not in todos_procedimentos_conhecidos:
                    mask = tabela[coluna_procedimento].astype(str) == str(procedimento)
                    quantidade = tabela.loc[mask, coluna_quantidade].sum()
                    
                    try:
                        quantidade_num = float(quantidade) if not pd.isna(quantidade) else 0
                    except (ValueError, TypeError):
                        quantidade_num = 0
                    
                    if quantidade_num > 0:
                        procedimentos_nao_mapeados[procedimento] = quantidade_num

            if procedimentos_nao_mapeados:
                procedimentos_nao_mapeados = dict(sorted(procedimentos_nao_mapeados.items(), 
                                                        key=lambda x: x[1], reverse=True))
                
                print(f"Encontrados {len(procedimentos_nao_mapeados)} procedimentos não mapeados em {municipio}")
                
                # Criar DataFrame para não listados deste município
                df_nao_listado = pd.DataFrame({
                    f"Procedimento_{municipio}": list(procedimentos_nao_mapeados.keys()),
                    f"Quantidade_{municipio}": list(procedimentos_nao_mapeados.values())
                })
                nao_listados_dfs.append(df_nao_listado)
            
        except FileNotFoundError:
            print(f"Arquivo {arquivo} não encontrado!")
            continue  

        except Exception as e:
            print(f"Erro em {municipio}: {e}")
            continue

    if resultados_por_municipio: 
        todos_procedimentos = set()
        for municipio, totais in resultados_por_municipio.items():
            todos_procedimentos.update(totais.keys())
        
        todos_procedimentos = sorted(list(todos_procedimentos))
        
        df_consolidado = pd.DataFrame(index=todos_procedimentos)
        
        # Adicionar colunas para cada município
        for municipio in municipios:
            totais = resultados_por_municipio.get(municipio, {})

            valores = []
            for procedimento in todos_procedimentos:
                valores.append(totais.get(procedimento, 0))
            
            df_consolidado[municipio] = valores
        
        # Adicionar linha de totais
        df_consolidado.loc['TOTAL'] = df_consolidado.sum()
        
        # Salvar em Excel
        df_consolidado.to_excel("Prestador/neomater/resultado/TODOS_MUNICIPIOS_CONSOLIDADO.xlsx")
        print(f"\nArquivo consolidado salvo com {len(df_consolidado.columns)} colunas")
    
    # CRIAR EXCEL COM PROCEDIMENTOS NÃO LISTADOS
    if nao_listados_dfs:
        # Juntar horizontalmente todos os DataFrames de não listados
        if len(nao_listados_dfs) > 1:
            df_nao_listados_consolidado = pd.concat(nao_listados_dfs, axis=1)
        else:
            df_nao_listados_consolidado = nao_listados_dfs[0]
        
        df_nao_listados_consolidado.to_excel("Prestador/neomater/resultado/PROCEDIMENTOS_NAO_LISTADOS_CONSOLIDADO.xlsx", index=False)
        print(f"Arquivo de procedimentos não listados salvo com {len(df_nao_listados_consolidado.columns)} colunas")
