"""
SCRIPT PARA PROCESSAR EXCEL E GERAR JSON CONSOLIDADO
Autor: Assistente IA
Data: 2024
"""

import pandas as pd
import numpy as np
import json
import os
import time
from datetime import datetime
from collections import defaultdict
import gc

def configurar_pandas():
    """Configura pandas para melhor performance"""
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.width', 1000)
    pd.set_option('display.max_colwidth', 50)

def ler_planilha_com_filtros(caminho_excel, nome_planilha, usecols=None, nrows=None, dtype=None):
    """
    Lê planilha Excel com otimizações de memória
    """
    print(f"Lendo planilha: {nome_planilha}")
    
    try:
        # Ler com otimizações
        df = pd.read_excel(
            caminho_excel,
            sheet_name=nome_planilha,
            usecols=usecols,  # Apenas colunas necessárias
            nrows=nrows,      # Limitar linhas para teste
            dtype=dtype,      # Tipos de dados otimizados
            engine='openpyxl'
        )
        
        print(f"  ✓ Linhas lidas: {len(df):,}")
        print(f"  ✓ Colunas: {list(df.columns)}")
        
        # Remover espaços em branco nas strings
        str_cols = df.select_dtypes(include=['object']).columns
        for col in str_cols:
            df[col] = df[col].astype(str).str.strip()
        
        return df
    
    except Exception as e:
        print(f"  ✗ Erro ao ler {nome_planilha}: {e}")
        return None

def processar_base_pedidos(df_pedidos):
    """
    Processa base_pedidos e retorna estrutura consolidada
    """
    print("\n" + "="*60)
    print("PROCESSANDO BASE_PEDIDOS")
    print("="*60)
    
    if df_pedidos is None or df_pedidos.empty:
        print("Nenhum dado para processar")
        return {}
    
    # Converter colunas numéricas
    colunas_numericas = ['ULT_STATUS', 'PROX_STATUS', 'QTD', 'CLIENTE', 'MCCO']
    for col in colunas_numericas:
        if col in df_pedidos.columns:
            df_pedidos[col] = pd.to_numeric(df_pedidos[col], errors='coerce')
    
    # Filtrar apenas status relevantes
    mask_relevante = (
        ((df_pedidos['ULT_STATUS'] == 520) & (df_pedidos['PROX_STATUS'] == 540)) |
        ((df_pedidos['ULT_STATUS'] == 541) & (df_pedidos['PROX_STATUS'] == 540)) |
        ((df_pedidos['ULT_STATUS'] == 980) & (df_pedidos['PROX_STATUS'] == 999)) |
        ((df_pedidos['ULT_STATUS'] == 982) & (df_pedidos['PROX_STATUS'] == 999)) |
        ((df_pedidos['ULT_STATUS'] == 983) & (df_pedidos['PROX_STATUS'] == 999))
    )
    
    df_relevante = df_pedidos[mask_relevante].copy()
    print(f"Pedidos relevantes encontrados: {len(df_relevante):,} de {len(df_pedidos):,}")
    
    if df_relevante.empty:
        return {}
    
    # Converter datas
    if 'DT_TRANS' in df_relevante.columns:
        df_relevante['DT_TRANS'] = pd.to_datetime(df_relevante['DT_TRANS'], errors='coerce')
    
    # Criar chave única
    df_relevante['CHAVE'] = df_relevante['NUM_PEDIDO_SISCAP'] + '|' + df_relevante['CLIENTE'].astype(str)
    
    # Estrutura para armazenar resultados
    resultados = {}
    
    # Processar por grupo (chave única)
    total_grupos = df_relevante['CHAVE'].nunique()
    print(f"Grupos únicos a processar: {total_grupos:,}")
    
    for i, (chave, grupo) in enumerate(df_relevante.groupby('CHAVE'), 1):
        # Progresso
        if i % 100 == 0:
            print(f"  Processando grupo {i:,} de {total_grupos:,}")
        
        # Construir estrutura para este grupo
        estrutura = construir_estrutura_grupo(grupo)
        
        if estrutura:
            resultados[chave] = estrutura
    
    print(f"\n✓ Processados {len(resultados):,} grupos com sucesso")
    return resultados

def construir_estrutura_grupo(grupo):
    """
    Constrói estrutura JSON para um grupo de pedidos
    """
    try:
        # Dados do cabeçalho (primeira linha do grupo)
        primeira_linha = grupo.iloc[0]
        
        estrutura = {
            'NUM_PEDIDO_SISCAP': str(primeira_linha.get('NUM_PEDIDO_SISCAP', '')),
            'CLIENTE': str(primeira_linha.get('CLIENTE', '')),
            'NOME_CLIENTE': str(primeira_linha.get('NOME_CLI', '')),
            'MCCO': str(primeira_linha.get('MCCO', '')),
            'PEDIDO_NOVO': [],
            'PEDIDO_WMS': [],  # Será preenchido depois
            'PEDIDO_POSTADO': []  # Será preenchido depois
        }
        
        # Agrupar por PED e TP_PED
        grupo_por_pedido = grupo.groupby(['PED', 'TP_PED'])
        
        for (pedido_num, tipo_pedido), subgrupo in grupo_por_pedido:
            # Primeira linha do subgrupo
            primeira_sub = subgrupo.iloc[0]
            
            pedido_novo = {
                'DATA_PEDIDO': primeira_sub['DT_TRANS'].strftime('%Y-%m-%d') if pd.notna(primeira_sub['DT_TRANS']) else '',
                'PEDIDO': str(pedido_num),
                'TIPO_PEDIDO': str(tipo_pedido),
                'CLIENTE': str(primeira_sub.get('CLIENTE', '')),
                'RESERVADO': [],
                'CANCELADO': []
            }
            
            # Separar itens por status
            # RESERVADO: 520→540 ou 541→540
            mask_reservado = (
                ((subgrupo['ULT_STATUS'] == 520) & (subgrupo['PROX_STATUS'] == 540)) |
                ((subgrupo['ULT_STATUS'] == 541) & (subgrupo['PROX_STATUS'] == 540))
            )
            
            # CANCELADO: 980/982/983→999
            mask_cancelado = (
                ((subgrupo['ULT_STATUS'] == 980) & (subgrupo['PROX_STATUS'] == 999)) |
                ((subgrupo['ULT_STATUS'] == 982) & (subgrupo['PROX_STATUS'] == 999)) |
                ((subgrupo['ULT_STATUS'] == 983) & (subgrupo['PROX_STATUS'] == 999))
            )
            
            # Processar itens RESERVADOS
            itens_reservados = subgrupo[mask_reservado]
            for _, item in itens_reservados.iterrows():
                pedido_novo['RESERVADO'].append({
                    'CODIGO': str(item.get('ITEM', '')),
                    'DESCRICAO_CODIGO': str(item.get('DESC_ITEM', '')),
                    'UM': str(item.get('UN_MEDIDA', '')),
                    'QTDE': str(int(item.get('QTD', 0))) if pd.notna(item.get('QTD')) else '0'
                })
            
            # Processar itens CANCELADOS
            itens_cancelados = subgrupo[mask_cancelado]
            for _, item in itens_cancelados.iterrows():
                pedido_novo['CANCELADO'].append({
                    'CODIGO': str(item.get('ITEM', '')),
                    'DESCRICAO_CODIGO': str(item.get('DESC_ITEM', '')),
                    'UM': str(item.get('UN_MEDIDA', '')),
                    'QTDE': str(int(item.get('QTD', 0))) if pd.notna(item.get('QTD')) else '0'
                })
            
            estrutura['PEDIDO_NOVO'].append(pedido_novo)
        
        return estrutura
        
    except Exception as e:
        print(f"Erro ao construir grupo: {e}")
        return None

def processar_outras_bases(df_lotes, df_postados, resultados_pedidos):
    """
    Processa base_lotes e base_objeto_postado
    """
    print("\n" + "="*60)
    print("PROCESSANDO BASES LOTE E POSTADO")
    print("="*60)
    
    # TODO: Implementar lógica para integrar estas bases
    # Por enquanto, retorna resultados sem estas bases
    print("⚠️  Funcionalidade em desenvolvimento")
    
    return resultados_pedidos

def salvar_resultados(resultados, caminho_saida):
    """
    Salva resultados em arquivos JSON
    """
    print("\n" + "="*60)
    print("SALVANDO RESULTADOS")
    print("="*60)
    
    if not resultados:
        print("Nenhum resultado para salvar")
        return
    
    # Criar pasta de saída se não existir
    os.makedirs(caminho_saida, exist_ok=True)
    
    total_salvo = 0
    
    for chave, estrutura in resultados.items():
        try:
            # Gerar nome de arquivo seguro
            nome_arquivo = f"pedido_{chave.replace('|', '_')}.json"
            caminho_completo = os.path.join(caminho_saida, nome_arquivo)
            
            # Salvar JSON
            with open(caminho_completo, 'w', encoding='utf-8') as f:
                json.dump(estrutura, f, ensure_ascii=False, indent=2)
            
            total_salvo += 1
            
        except Exception as e:
            print(f"Erro ao salvar {chave}: {e}")
    
    print(f"\n✓ Salvos {total_salvo} arquivos JSON em: {caminho_saida}")
    
    # Também salvar resumo em Excel
    salvar_resumo_excel(resultados, caminho_saida)

def salvar_resumo_excel(resultados, caminho_saida):
    """
    Cria um resumo em Excel com os dados processados
    """
    resumo_data = []
    
    for chave, estrutura in resultados.items():
        for pedido in estrutura.get('PEDIDO_NOVO', []):
            resumo_data.append({
                'NUM_PEDIDO_SISCAP': estrutura['NUM_PEDIDO_SISCAP'],
                'CLIENTE': estrutura['CLIENTE'],
                'NOME_CLIENTE': estrutura['NOME_CLIENTE'],
                'MCCO': estrutura['MCCO'],
                'PEDIDO': pedido['PEDIDO'],
                'TIPO_PEDIDO': pedido['TIPO_PEDIDO'],
                'DATA_PEDIDO': pedido['DATA_PEDIDO'],
                'QTD_RESERVADOS': len(pedido['RESERVADO']),
                'QTD_CANCELADOS': len(pedido['CANCELADO'])
            })
    
    if resumo_data:
        df_resumo = pd.DataFrame(resumo_data)
        caminho_excel = os.path.join(caminho_saida, 'resumo_processamento.xlsx')
        df_resumo.to_excel(caminho_excel, index=False)
        print(f"✓ Resumo salvo em Excel: {caminho_excel}")

def testar_com_pequeno_conjunto(caminho_excel, nrows=1000):
    """
    Testa o processamento com um conjunto pequeno de dados
    """
    print("🔬 TESTE COM CONJUNTO PEQUENO")
    print("="*60)
    
    # Ler apenas algumas linhas para teste
    df_pedidos = ler_planilha_com_filtros(
        caminho_excel,
        'base_pedidos',
        nrows=nrows,
        dtype={
            'NUM_PEDIDO_SISCAP': str,
            'CLIENTE': str,
            'NOME_CLI': str,
            'MCCO': str,
            'PED': str,
            'TP_PED': str,
            'ITEM': str,
            'DESC_ITEM': str,
            'UN_MEDIDA': str,
            'ULT_STATUS': 'int32',
            'PROX_STATUS': 'int32',
            'QTD': 'int32'
        }
    )
    
    if df_pedidos is not None:
        resultados = processar_base_pedidos(df_pedidos)
        
        if resultados:
            # Mostrar exemplo do primeiro resultado
            primeira_chave = list(resultados.keys())[0]
            print(f"\n📄 EXEMPLO DO PRIMEIRO JSON ({primeira_chave}):")
            print(json.dumps(resultados[primeira_chave], indent=2, ensure_ascii=False)[:500] + "...")
            
            return True
    return False

def main():
    """
    Função principal
    """
    print("="*70)
    print("PROCESSADOR DE EXCEL PARA JSON")
    print("="*70)
    
    # Configurar
    configurar_pandas()
    
    # Caminhos (ajuste conforme necessário)
    #caminho_excel = "base-completa-teste.xlsx"  # Ou caminho completo
    caminho_excel = r"c:\app\dashboard\001-base_completa.xlsm"  # Ou caminho completo
    caminho_saida = "resultados_json"
    
    # Verificar se arquivo existe
    if not os.path.exists(caminho_excel):
        print(f"✗ Arquivo não encontrado: {caminho_excel}")
        print("Por favor, ajuste o caminho do arquivo Excel.")
        return
    
    # Teste inicial com poucos dados
    print("\n🎯 FASE 1: TESTE INICIAL (100 registros)")
    if not testar_com_pequeno_conjunto(caminho_excel, nrows=100):
        print("Teste inicial falhou. Verifique o arquivo Excel.")
        return
    
    # Perguntar ao usuário
    print("\n" + "="*60)
    print("OPÇÕES DE PROCESSAMENTO:")
    print("1. Processar apenas 1.000 registros (teste rápido)")
    print("2. Processar 10.000 registros (teste médio)")
    print("3. Processar TODOS os registros")
    print("4. Processar quantidade personalizada")
    
    try:
        opcao = input("\nDigite sua opção (1-4): ").strip()
        
        if opcao == '1':
            limite = 1000
        elif opcao == '2':
            limite = 200000
        elif opcao == '3':
            limite = None  # Todos
        elif opcao == '4':
            limite = int(input("Digite a quantidade: "))
        else:
            print("Opção inválida. Usando padrão (1.000 registros).")
            limite = 1000
        
    except:
        print("Entrada inválida. Usando padrão (1.000 registros).")
        limite = 1000
    
    # Processar
    print(f"\n🚀 FASE 2: PROCESSAMENTO PRINCIPAL ({'TODOS' if limite is None else limite:,} registros)")
    
    inicio_total = time.time()
    
    # Ler dados com limite
    df_pedidos = ler_planilha_com_filtros(
        caminho_excel,
        'base_pedidos',
        nrows=limite,
        dtype={
            'NUM_PEDIDO_SISCAP': str,
            'CLIENTE': 'int32',
            'NOME_CLI': str,
            'MCCO': 'int32',
            'PED': 'int32',
            'TP_PED': str,
            'ITEM': str,
            'DESC_ITEM': str,
            'UN_MEDIDA': str,
            'ULT_STATUS': 'int32',
            'PROX_STATUS': 'int32',
            'QTD': 'int32'
        }
    )
    
    if df_pedidos is None:
        print("Falha ao ler dados.")
        return
    
    # Processar
    resultados = processar_base_pedidos(df_pedidos)
    
    # Salvar resultados
    if resultados:
        salvar_resultados(resultados, caminho_saida)
    
    # Estatísticas finais
    fim_total = time.time()
    tempo_total = fim_total - inicio_total
    
    print("\n" + "="*70)
    print("RESUMO FINAL DO PROCESSAMENTO")
    print("="*70)
    print(f"📊 Total de grupos processados: {len(resultados):,}")
    
    total_pedidos = sum(len(grupo['PEDIDO_NOVO']) for grupo in resultados.values())
    total_itens = sum(
        len(pedido['RESERVADO']) + len(pedido['CANCELADO'])
        for grupo in resultados.values()
        for pedido in grupo['PEDIDO_NOVO']
    )
    
    print(f"📊 Total de pedidos únicos: {total_pedidos:,}")
    print(f"📊 Total de itens processados: {total_itens:,}")
    print(f"⏱️  Tempo total: {tempo_total:.2f} segundos")
    
    if limite:
        velocidade = limite / tempo_total
        print(f"⚡ Velocidade: {velocidade:.0f} registros/segundo")
    
    print(f"\n✅ Processamento concluído!")
    print(f"📁 Resultados salvos em: {os.path.abspath(caminho_saida)}")
    print("="*70)

if __name__ == "__main__":
    main()