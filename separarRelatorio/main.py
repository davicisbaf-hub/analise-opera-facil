from openpyxl import load_workbook
from dotenv import load_dotenv
import pandas as pd
import os
import re

load_dotenv()

# Função para criar planilha "Por Município Colunas" (modificada)
def criar_planilha_municipio_colunas(caminho_arquivo):
    """
    Cria uma planilha onde cada município tem 3 colunas:
    1. Paciente [municipio]
    2. [municipio]  (apenas o nome do município)
    3. Quantidade [municipio]
    """
    
    wb = load_workbook(caminho_arquivo)
    ws = wb.active
    
    # Obter células mescladas
    merged_ranges = ws.merged_cells.ranges
    
    merged_cells_dict = {}
    for merged_range in merged_ranges:
        min_col, min_row, max_col, max_row = merged_range.bounds
        valor = ws.cell(row=min_row, column=min_col).value
        
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                merged_cells_dict[(row, col)] = valor
    
    # Coletar dados organizados por município
    dados_por_municipio = {}
    
    municipio_atual = None
    estado_atual = None
    
    for row in range(1, ws.max_row + 1):
        # Verificar se é linha de município
        cell_value = str(ws.cell(row=row, column=1).value) if ws.cell(row=row, column=1).value else ""
        
        if 'MUNICIPIO:' in cell_value:
            if '-' in cell_value:
                partes = cell_value.split('-')
                estado_atual = partes[0].replace('MUNICIPIO:', '').strip()
                municipio_atual = f"{estado_atual} - {partes[1].strip()}"
            continue
        
        # Verificar se é linha de dados
        data_hora = None
        if (row, 1) in merged_cells_dict:
            data_hora = merged_cells_dict[(row, 1)]
        else:
            data_hora = ws.cell(row=row, column=1).value
        
        if data_hora and isinstance(data_hora, str) and re.match(r'\d{2}/\d{2}/\d{4}', data_hora[:10]):
            # Extrair dados
            linha_dados = []
            for col in range(1, 11):
                if (row, col) in merged_cells_dict:
                    valor = merged_cells_dict[(row, col)]
                else:
                    valor = ws.cell(row=row, column=col).value
                linha_dados.append(valor)
            
            if len(linha_dados) >= 10 and municipio_atual:
                paciente = linha_dados[1]
                procedimento = linha_dados[3]
                quantidade = linha_dados[4]
                
                if paciente and procedimento:
                    try:
                        qtd = int(float(quantidade)) if quantidade else 1
                    except:
                        qtd = 1
                    
                    # Adicionar aos dados do município
                    if municipio_atual not in dados_por_municipio:
                        dados_por_municipio[municipio_atual] = []
                    
                    dados_por_municipio[municipio_atual].append({
                        'paciente': paciente,
                        'procedimento': procedimento,
                        'quantidade': qtd
                    })
    
    # Se não encontrou dados, retornar None
    if not dados_por_municipio:
        print(f"⚠️ Nenhum dado encontrado para organizar por município")
        return None, []
    
    # Preparar dados para o DataFrame
    municipios = list(dados_por_municipio.keys())
    
    # Encontrar o máximo de registros em um município
    max_registros = max(len(dados) for dados in dados_por_municipio.values())
    
    # Criar dicionário para o DataFrame
    df_data = {}
    
    # Para cada município, criar as 3 colunas
    for municipio in municipios:
        dados = dados_por_municipio[municipio]
        
        # Criar listas para as 3 colunas
        pacientes = []
        procedimentos = []
        quantidades = []
        
        # Preencher com dados
        for dado in dados:
            pacientes.append(dado['paciente'])
            procedimentos.append(dado['procedimento'])
            quantidades.append(dado['quantidade'])
        
        # Preencher com None para alinhar com outros municípios
        while len(pacientes) < max_registros:
            pacientes.append(None)
            procedimentos.append(None)
            quantidades.append(None)
        
        # Adicionar ao dicionário do DataFrame
        # MODIFICAÇÃO AQUI: Apenas o nome do município na coluna do meio
        df_data[f'Paciente {municipio}'] = pacientes
        df_data[f'{municipio}'] = procedimentos  # Apenas o nome do município
        df_data[f'Quantidade {municipio}'] = quantidades
    
    # Criar DataFrame
    df_final = pd.DataFrame(df_data)
    
    # Remover linhas totalmente vazias
    df_final = df_final.dropna(how='all')
    
    # Reorganizar as colunas para agrupar por município
    colunas_ordenadas = []
    for municipio in municipios:
        colunas_ordenadas.append(f'Paciente {municipio}')
        colunas_ordenadas.append(f'{municipio}')  # Apenas o nome do município
        colunas_ordenadas.append(f'Quantidade {municipio}')
    
    df_final = df_final[colunas_ordenadas]
    
    return df_final, municipios

# Função para criar planilha "Dados Detalhados" (mantida igual)
def criar_planilha_dados_detalhados(caminho_arquivo):
    """
    Cria uma planilha com todos os dados detalhados, incluindo município.
    """
    
    wb = load_workbook(caminho_arquivo)
    ws = wb.active
    
    # Obter células mescladas
    merged_ranges = ws.merged_cells.ranges
    
    merged_cells_dict = {}
    for merged_range in merged_ranges:
        min_col, min_row, max_col, max_row = merged_range.bounds
        valor = ws.cell(row=min_row, column=min_col).value
        
        for row in range(min_row, max_row + 1):
            for col in range(min_col, max_col + 1):
                merged_cells_dict[(row, col)] = valor
    
    # Coletar dados detalhados
    dados_detalhados = []
    municipio_atual = None
    estado_atual = None
    
    for row in range(1, ws.max_row + 1):
        # Verificar se é linha de município
        cell_value = str(ws.cell(row=row, column=1).value) if ws.cell(row=row, column=1).value else ""
        
        if 'MUNICIPIO:' in cell_value:
            if '-' in cell_value:
                partes = cell_value.split('-')
                estado_atual = partes[0].replace('MUNICIPIO:', '').strip()
                municipio_atual = f"{estado_atual} - {partes[1].strip()}"
            continue
        
        # Verificar se é linha de dados
        data_hora = None
        if (row, 1) in merged_cells_dict:
            data_hora = merged_cells_dict[(row, 1)]
        else:
            data_hora = ws.cell(row=row, column=1).value
        
        if data_hora and isinstance(data_hora, str) and re.match(r'\d{2}/\d{2}/\d{4}', data_hora[:10]):
            # Extrair dados
            linha_dados = []
            for col in range(1, 11):
                if (row, col) in merged_cells_dict:
                    valor = merged_cells_dict[(row, col)]
                else:
                    valor = ws.cell(row=row, column=col).value
                linha_dados.append(valor)
            
            if len(linha_dados) >= 10 and municipio_atual:
                # Adicionar dados com município
                linha_completa = [municipio_atual] + linha_dados
                dados_detalhados.append(linha_completa)
    
    # Criar DataFrame
    if dados_detalhados:
        colunas = ['Município', 'Data/Hora', 'Paciente', 'Data Nascimento', 
                  'Procedimento', 'Quantidade', 'Valor Regional', 'Contraste', 
                  'Sedação', 'Valor SUS', 'Valor Total']
        
        # Ajustar colunas se necessário
        if len(dados_detalhados[0]) < len(colunas):
            colunas = colunas[:len(dados_detalhados[0])]
        
        df_detalhado = pd.DataFrame(dados_detalhados, columns=colunas)
        return df_detalhado
    else:
        return None

# Função principal simplificada (atualizada)
def processar_relatorio_simplificado(caminho_arquivo):
    """
    Processa o relatório e salva apenas as 2 planilhas solicitadas:
    1. Por Município Colunas (com formato modificado)
    2. Dados Detalhados
    """
    
    print(f"📊 Processando: {caminho_arquivo}")
    
    # Criar as duas planilhas
    df_municipio_colunas, municipios_encontrados = criar_planilha_municipio_colunas(caminho_arquivo)
    df_dados_detalhados = criar_planilha_dados_detalhados(caminho_arquivo)
    
    # Verificar se temos dados
    if df_municipio_colunas is None and df_dados_detalhados is None:
        print(f"❌ Nenhum dado encontrado no arquivo: {caminho_arquivo}")
        return None
    
    # Criar diretório de saída
    output_dir = "relatorios_simplificados"
    if not os.path.exists(output_dir):
        os.makedirs(output_dir)
    
    # Nome do arquivo base
    nome_base = os.path.splitext(os.path.basename(caminho_arquivo))[0]
    nome_arquivo = f'{output_dir}/{nome_base}_SIMPLIFICADO.xlsx'
    
    # Salvar em Excel com as 2 abas
    with pd.ExcelWriter(nome_arquivo, engine='openpyxl') as writer:
        # Aba 1: Por Município Colunas (FORMATO MODIFICADO)
        if df_municipio_colunas is not None:
            df_municipio_colunas.to_excel(writer, sheet_name='Por Município Colunas', index=False)
            print(f"   ✅ Planilha 'Por Município Colunas' criada (formato modificado)")
        
        # Aba 2: Dados Detalhados
        if df_dados_detalhados is not None:
            df_dados_detalhados.to_excel(writer, sheet_name='Dados Detalhados', index=False)
            print(f"   ✅ Planilha 'Dados Detalhados' criada")
    
    print(f"✅ Arquivo salvo: {nome_arquivo}")
    
    # Mostrar resumo
    print(f"\n📋 RESUMO DO ARQUIVO:")
    print("-" * 60)
    
    if df_municipio_colunas is not None:
        print(f"• Planilha 'Por Município Colunas':")
        print(f"  - Municípios: {len(municipios_encontrados)}")
        print(f"  - Colunas: {len(df_municipio_colunas.columns)}")
        print(f"  - Linhas: {len(df_municipio_colunas)}")
    
    if df_dados_detalhados is not None:
        print(f"• Planilha 'Dados Detalhados':")
        print(f"  - Registros: {len(df_dados_detalhados)}")
        print(f"  - Colunas: {len(df_dados_detalhados.columns)}")
    
    # Mostrar exemplo do NOVO formato
    if df_municipio_colunas is not None and len(municipios_encontrados) > 0:
        print(f"\n📋 EXEMPLO DO NOVO FORMATO 'Por Município Colunas':")
        print("=" * 100)
        
        # Mostrar estrutura das colunas (NOVO FORMATO)
        print(f"\nEstrutura das colunas (primeiros 2 municípios como exemplo):")
        for municipio in municipios_encontrados[:2]:
            print(f"  • Paciente {municipio}")
            print(f"  • {municipio}")  # Apenas o nome do município
            print(f"  • Quantidade {municipio}")
        
        # Mostrar algumas linhas de exemplo
        if len(df_municipio_colunas) > 0:
            print(f"\nPrimeiras 2 linhas de exemplo:")
            print("-" * 80)
            
            for i in range(min(2, len(df_municipio_colunas))):
                print(f"Linha {i+1}:")
                
                # Mostrar apenas dados não nulos
                for municipio in municipios_encontrados[:2]:  # Apenas 2 municípios como exemplo
                    paciente_col = f'Paciente {municipio}'
                    
                    if paciente_col in df_municipio_colunas.columns:
                        paciente = df_municipio_colunas.iloc[i][paciente_col]
                        if pd.notna(paciente):
                            print(f"  Município: {municipio}")
                            print(f"    Paciente: {paciente}")
                            print(f"    Procedimento: {df_municipio_colunas.iloc[i][f'{municipio}']}")
                            print(f"    Quantidade: {df_municipio_colunas.iloc[i][f'Quantidade {municipio}']}")
                print("-" * 40)
    
    return {
        'df_municipio_colunas': df_municipio_colunas,
        'df_dados_detalhados': df_dados_detalhados,
        'municipios_encontrados': municipios_encontrados,
        'arquivo_saida': nome_arquivo
    }

# Função para processar todos os arquivos (mantida igual)
def processar_todos_arquivos_simplificado():
    """
    Processa todos os arquivos listados na variável de ambiente
    mantendo apenas as 2 planilhas solicitadas.
    """
    # Ler a variável de ambiente
    arquivos_str = os.getenv("separarArquivo", "")
    
    if not arquivos_str:
        print("⚠️ Variável de ambiente 'separarArquivo' não encontrada!")
        # Usar arquivos padrão se não encontrar
        arquivos = [
            "separarRelatorio/separarNeomater.xlsx",
            "separarRelatorio/separarNeotin.xlsx", 
            "separarRelatorio/separarPediatrico.xlsx"
        ]
    else:
        # Limpar e processar a string
        arquivos_str = arquivos_str.strip()
        if arquivos_str.startswith('["') and arquivos_str.endswith('"]'):
            arquivos_str = arquivos_str[2:-2]
        
        arquivos = [arquivo.strip() for arquivo in arquivos_str.split(',') if arquivo.strip()]
    
    print(f"📁 Total de arquivos a processar: {len(arquivos)}")
    
    resultados = []
    arquivos_processados = 0
    
    for arquivo in arquivos:
        print("\n" + "=" * 80)
        print(f"🚀 PROCESSANDO: {arquivo}")
        print("=" * 80)
        
        if not os.path.exists(arquivo):
            print(f"❌ Arquivo não encontrado: {arquivo}")
            continue
        
        try:
            resultado = processar_relatorio_simplificado(arquivo)
            if resultado is not None:
                resultados.append(resultado)
                arquivos_processados += 1
                print(f"✅ Concluído: {arquivo}")
            
        except Exception as e:
            print(f"❌ Erro ao processar {arquivo}: {str(e)}")
            import traceback
            traceback.print_exc()
    
    print("\n" + "=" * 80)
    print("🎉 PROCESSAMENTO CONCLUÍDO!")
    print("=" * 80)
    
    # Criar um resumo geral
    if resultados:
        print(f"\n📈 RESUMO GERAL:")
        print("-" * 60)
        
        total_municipios = 0
        total_registros = 0
        
        for resultado in resultados:
            arquivo = resultado['arquivo_saida']
            num_municipios = len(resultado['municipios_encontrados'])
            
            num_registros_detalhados = 0
            if resultado['df_dados_detalhados'] is not None:
                num_registros_detalhados = len(resultado['df_dados_detalhados'])
            
            print(f"📄 {os.path.basename(arquivo)}:")
            print(f"   • Municípios encontrados: {num_municipios}")
            print(f"   • Registros detalhados: {num_registros_detalhados}")
            
            total_municipios += num_municipios
            total_registros += num_registros_detalhados
        
        print(f"\n📊 TOTAIS:")
        print(f"   • Arquivos processados: {arquivos_processados}/{len(arquivos)}")
        print(f"   • Total de municípios distintos: {total_municipios}")
        print(f"   • Total de registros detalhados: {total_registros}")
        
        # Mostrar exemplo da estrutura de arquivo final
        if resultados:
            print(f"\n📂 ESTRUTURA DOS ARQUIVOS GERADOS:")
            print("-" * 60)
            print(f"Todos os arquivos foram salvos na pasta: relatorios_simplificados/")
            print(f"Cada arquivo contém 2 planilhas:")
            print(f"  1. 'Por Município Colunas' - Formato com colunas por município")
            print(f"      • Paciente [município]")
            print(f"      • [município] (apenas o nome)")
            print(f"      • Quantidade [município]")
            print(f"  2. 'Dados Detalhados' - Dados completos organizados")
    
    return resultados

# Executar
if __name__ == "__main__":
    # Processar todos os arquivos
    resultados = processar_todos_arquivos_simplificado()
    
    # Mensagem final
    if resultados:
        print(f"\n✨ Todos os relatórios foram processados com sucesso!")
        print(f"📍 Os arquivos estão na pasta: relatorios_simplificados/")