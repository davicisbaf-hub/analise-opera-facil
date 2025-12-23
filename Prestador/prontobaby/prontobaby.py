import pandas as pd

from procedimentos import (
    pacote_otorrino,
    pacote_geral,
    pacote_oftalmo,
    pacote_hispospadia,
    pacote_inguinal,
    pacote_hidrocele,
    pacote_adeno,
    pacote_amig,
    pacote_amig_adeno,
    pacote_estrabismo,
    pacote_nasal,
    pacote_orqui,
    pacote_plastica,
    pacote_postec,
    pacote_umbilical
)

def analisar_prontobaby():
    arquivo = "relatorios_simplificados/separarPediatrico_SIMPLIFICADO.xlsx"
    municipios = ["RJ - Belford Roxo", "RJ - Duque de Caxias","RJ - Itaguaí", "RJ - Japeri", "RJ - Magé", "RJ - Mesquita", "RJ - Nilópolis", "RJ - Nova Iguaçu", "RJ - Paracambi", "RJ - Queimados", "RJ - Seropédica", "RJ - São João de Meriti"]

    for municipio in municipios:
        try:
            # carregar listas de procedimentos
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
            coluna_quantidade = "Quantidade {}".format(municipio)

            if coluna_procedimento not in tabela.columns:
                print(f"Coluna '{coluna_procedimento}' não encontrada!")
                continue
            
            if coluna_quantidade not in tabela.columns:
                print(f"Coluna '{coluna_quantidade}' não encontrada!")
                continue
            
            resultados = {}
            
            def process_group(procedimentos):
                total = 0
                for procedimento in procedimentos:
                    mask = tabela[coluna_procedimento].astype(str) == str(procedimento)
                    quantidade = tabela.loc[mask, coluna_quantidade].sum()
                    
                    # CORREÇÃO: Converter para número e verificar se é válido
                    try:
                        quantidade_num = float(quantidade) if not pd.isna(quantidade) else 0
                    except (ValueError, TypeError):
                        quantidade_num = 0
                    
                    # CORREÇÃO: Verificar se é maior que 0
                    if quantidade_num > 0:
                        resultados[procedimento] = quantidade_num
                        total += quantidade_num
                return total
            
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
            
            # CORREÇÃO: Criar DataFrame corretamente
            if resultados:
                for nome, lista in grupos.items():
                    totais[nome] = process_group(lista or [])
                    procedimentoMunicipio = {
                        coluna_procedimento: totais.keys(),
                        coluna_quantidade: totais.values(),
                    }


                df = pd.DataFrame(procedimentoMunicipio)
                df.to_excel(f"Prestador/prontobaby/resultado/{coluna_procedimento}.xlsx", index=False)

            print("=== RESULTADOS {} ===".format(coluna_procedimento))
            for nome, total in totais.items():
                print(f"{nome.replace('_', ' ').title()} Total: {total}")
            soma_total = sum(totais.values())
            print(f"Soma Total: {soma_total}")
            
            print("\n")
            
            todos_procedimentos_conhecidos = []
            for lista in grupos.values():
                todos_procedimentos_conhecidos.extend(lista or [])
            
            todos_procedimentos_conhecidos = list(set([str(p) for p in todos_procedimentos_conhecidos if p and str(p).strip()]))
            
            procedimentos_na_coluna = tabela[coluna_procedimento].dropna().unique()
            procedimentos_na_coluna = [str(p) for p in procedimentos_na_coluna if p and str(p).strip()]
            
            # Encontrar procedimentos que estão na coluna mas não nas listas
            procedimentos_nao_mapeados = {}
            for procedimento in procedimentos_na_coluna:
                if procedimento not in todos_procedimentos_conhecidos:
                    mask = tabela[coluna_procedimento].astype(str) == str(procedimento)
                    quantidade = tabela.loc[mask, coluna_quantidade].sum()
                    
                    # CORREÇÃO: Converter para número
                    try:
                        quantidade_num = float(quantidade) if not pd.isna(quantidade) else 0
                    except (ValueError, TypeError):
                        quantidade_num = 0
                    
                    if quantidade_num > 0:
                        procedimentos_nao_mapeados[procedimento] = quantidade_num

            if procedimentos_nao_mapeados:
                procedimentos_nao_mapeados = dict(sorted(procedimentos_nao_mapeados.items(), key=lambda x: x[1], reverse=True))
                
                print(f"\nEncontrados {len(procedimentos_nao_mapeados)} procedimentos não mapeados:")
                total_nao_mapeado = 0
                for procedimento, quantidade in procedimentos_nao_mapeados.items():
                    print(f"  '{procedimento}': {quantidade}")
                    total_nao_mapeado += quantidade

                # CORREÇÃO: Criar DataFrame para não mapeados
                procedimentoNaoListado = {
                    "Procedimento Nao listados": list(procedimentos_nao_mapeados.keys()),
                    "Quantidade Nao listados": list(procedimentos_nao_mapeados.values()),
                }

                df = pd.DataFrame(procedimentoNaoListado)
                df.to_excel(f"Prestador/prontobaby/resultado/Nao-listados_{coluna_procedimento}.xlsx", index=False)
            
        except FileNotFoundError:
            print(f"Arquivo {arquivo} não encontrado!")
            continue  # CORREÇÃO: continuar para o próximo município
        except Exception as e:
            print(f"Erro em {municipio}: {e}")
            continue  # CORREÇÃO: continuar para o próximo município