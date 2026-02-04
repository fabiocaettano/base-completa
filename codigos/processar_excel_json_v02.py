"""
PROCESSADOR COMPLETO DE EXCEL PARA JSON - VERSÃO CORRIGIDA
Processa base_pedidos, base_lotes e base_objeto_postado
Gera JSON estruturado com relacionamentos entre as bases
"""

import pandas as pd
import numpy as np
import json
import os
import time
from datetime import datetime
from collections import defaultdict
import gc
import warnings
warnings.filterwarnings('ignore')

def configurar_pandas():
    """Configura pandas para melhor performance"""
    pd.set_option('display.max_columns', None)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.width', 1000)
    pd.set_option('display.max_colwidth', 50)

def formatar_numero(num):
    """Formata número com separadores de milhar"""
    if num is None:
        return "N/A"
    return f"{num:,}"

def ler_todas_planilhas(caminho_excel, limite=None):
    """
    Lê todas as planilhas do arquivo Excel
    """
    print("📂 LENDO TODAS AS PLANILHAS...")
    
    planilhas = {}
    
    try:
        # Ler base_pedidos
        print("  Lendo base_pedidos...")
        usecols_pedidos = [
            'DR', 'DT_TRANS', 'PED', 'TP_PED', 'NOTA_SERIE', 'TIPO_NOTA',
            'CLIENTE', 'MCCO', 'NOME_CLI', 'ITEM', 'DESC_ITEM', 'QTD',
            'UN_MEDIDA', 'ULT_STATUS', 'PROX_STATUS', 'NUM_PEDIDO_SISCAP'
        ]
        
        df_pedidos = pd.read_excel(
            caminho_excel,
            sheet_name='base_pedidos',
            usecols=usecols_pedidos,
            nrows=limite,
            dtype={
                'DR': 'int32',
                'PED': 'int32',
                'TP_PED': str,
                'NOTA_SERIE': str,
                'TIPO_NOTA': str,
                'CLIENTE': 'int32',
                'MCCO': 'int32',
                'NOME_CLI': str,
                'ITEM': str,
                'DESC_ITEM': str,
                'QTD': 'int32',
                'UN_MEDIDA': str,
                'ULT_STATUS': 'int32',
                'PROX_STATUS': 'int32',
                'NUM_PEDIDO_SISCAP': str
            },
            engine='openpyxl'
        )
        
        # Converter data
        if 'DT_TRANS' in df_pedidos.columns:
            df_pedidos['DT_TRANS'] = pd.to_datetime(df_pedidos['DT_TRANS'], errors='coerce')
        
        planilhas['pedidos'] = df_pedidos
        print(f"    ✓ {formatar_numero(len(df_pedidos))} registros lidos")
        
        # Ler base_lotes
        print("  Lendo base_lotes...")
        usecols_lotes = [
            'lote_pedido', 'item', 'descricao', 'qtde', 'um', 'status_lote',
            'pedido', 'tipo', 'zona', 'status_impressao', 'posto',
            'lote_consolidado', 'cep', 'roteirizador', 'cod_cliente',
            'descricao_cliente', 'se'
        ]
        
        df_lotes = pd.read_excel(
            caminho_excel,
            sheet_name='base_lotes',
            usecols=usecols_lotes,
            dtype={
                'lote_pedido': 'int32',
                'item': str,
                'descricao': str,
                'qtde': 'int32',
                'um': str,
                'status_lote': 'int32',
                'pedido': 'int32',
                'tipo': str,
                'zona': str,
                'status_impressao': str,
                'posto': str,
                'lote_consolidado': str,
                'cep': str,
                'roteirizador': str,
                'cod_cliente': 'int32',
                'descricao_cliente': str,
                'se': str
            },
            engine='openpyxl'
        )
        
        planilhas['lotes'] = df_lotes
        print(f"    ✓ {formatar_numero(len(df_lotes))} registros lidos")
        
        # Ler base_objeto_postado
        print("  Lendo base_objeto_postado...")
        usecols_postado = [
            'CLIENTE', 'NOME', 'LOTE', 'PEDIDO', 'TIPO_PEDIDO', 'SE',
            'ZONA', 'REGISTRO', 'NOTA_FISCAL', 'SERIE', 'DATA_EXPEDICAO',
            'DATA_CANCELAMENTO'
        ]
        
        df_postado = pd.read_excel(
            caminho_excel,
            sheet_name='base_objeto_postado',
            usecols=usecols_postado,
            dtype={
                'CLIENTE': 'int32',
                'NOME': str,
                'LOTE': 'int32',
                'PEDIDO': 'int32',
                'TIPO_PEDIDO': str,
                'SE': str,
                'ZONA': str,
                'REGISTRO': str,
                'NOTA_FISCAL': str,
                'SERIE': str,
                'DATA_EXPEDICAO': 'datetime64[ns]',
                'DATA_CANCELAMENTO': 'datetime64[ns]'
            },
            engine='openpyxl'
        )
        
        planilhas['postado'] = df_postado
        print(f"    ✓ {formatar_numero(len(df_postado))} registros lidos")
        
        return planilhas
        
    except Exception as e:
        print(f"❌ Erro ao ler planilhas: {e}")
        return None

def criar_mapa_relacionamentos(df_pedidos):
    """
    Cria mapa de relacionamentos para busca rápida
    Retorna: { (PED, TP_PED, CLIENTE): NUM_PEDIDO_SISCAP }
    """
    print("\n🔗 CRIANDO MAPA DE RELACIONAMENTOS...")
    
    mapa = {}
    
    # Criar chave única e agrupar
    df_pedidos['CHAVE'] = df_pedidos['NUM_PEDIDO_SISCAP'] + '|' + df_pedidos['CLIENTE'].astype(str)
    
    # Para cada grupo único (NUM_PEDIDO_SISCAP + CLIENTE)
    for chave, grupo in df_pedidos.groupby('CHAVE'):
        num_pedido_siscap = grupo['NUM_PEDIDO_SISCAP'].iloc[0]
        cliente = grupo['CLIENTE'].iloc[0]
        
        # Para cada pedido/tipo neste grupo
        for _, row in grupo.iterrows():
            pedido = row['PED']
            tipo_pedido = row['TP_PED']
            
            # Mapa 1: (PED, TP_PED, CLIENTE) -> NUM_PEDIDO_SISCAP
            mapa[(pedido, tipo_pedido, cliente)] = num_pedido_siscap
    
    print(f"  ✓ Mapa criado com {formatar_numero(len(mapa))} relacionamentos")
    return mapa

def processar_pedidos_novos(df_pedidos, mapa_relacionamentos):
    """
    Processa base_pedidos e cria estrutura principal
    """
    print("\n📋 PROCESSANDO PEDIDOS NOVOS...")
    
    # Filtrar apenas status relevantes
    mask_relevante = (
        ((df_pedidos['ULT_STATUS'] == 520) & (df_pedidos['PROX_STATUS'] == 540)) |
        ((df_pedidos['ULT_STATUS'] == 541) & (df_pedidos['PROX_STATUS'] == 540)) |
        ((df_pedidos['ULT_STATUS'] == 980) & (df_pedidos['PROX_STATUS'] == 999)) |
        ((df_pedidos['ULT_STATUS'] == 982) & (df_pedidos['PROX_STATUS'] == 999)) |
        ((df_pedidos['ULT_STATUS'] == 983) & (df_pedidos['PROX_STATUS'] == 999))
    )
    
    df_relevante = df_pedidos[mask_relevante].copy()
    print(f"  Pedidos relevantes: {formatar_numero(len(df_relevante))} de {formatar_numero(len(df_pedidos))}")
    
    if df_relevante.empty:
        return {}
    
    # Estrutura principal
    resultados = defaultdict(lambda: {
        'NUM_PEDIDO_SISCAP': '',
        'CLIENTE': '',
        'NOME_CLIENTE': '',
        'MCCO': '',
        'PEDIDO_NOVO': [],
        'PEDIDO_WMS': [],
        'PEDIDO_POSTADO': []
    })
    
    # Agrupar por chave única
    df_relevante['CHAVE'] = df_relevante['NUM_PEDIDO_SISCAP'] + '|' + df_relevante['CLIENTE'].astype(str)
    grupos = df_relevante.groupby('CHAVE')
    
    total_grupos = len(grupos)
    print(f"  Grupos únicos a processar: {formatar_numero(total_grupos)}")
    
    for i, (chave, grupo) in enumerate(grupos, 1):
        if i % 100 == 0:
            print(f"    Processando grupo {formatar_numero(i)} de {formatar_numero(total_grupos)}")
        
        # Dados do cabeçalho
        primeira = grupo.iloc[0]
        resultados[chave]['NUM_PEDIDO_SISCAP'] = str(primeira['NUM_PEDIDO_SISCAP'])
        resultados[chave]['CLIENTE'] = str(primeira['CLIENTE'])
        resultados[chave]['NOME_CLIENTE'] = str(primeira['NOME_CLI'])
        resultados[chave]['MCCO'] = str(primeira['MCCO'])
        
        # Processar pedidos neste grupo
        pedidos_agrupados = grupo.groupby(['PED', 'TP_PED'])
        
        for (pedido_num, tipo_pedido), subgrupo in pedidos_agrupados:
            pedido_novo = {
                'DATA_PEDIDO': subgrupo['DT_TRANS'].iloc[0].strftime('%Y-%m-%d') if pd.notna(subgrupo['DT_TRANS'].iloc[0]) else '',
                'PEDIDO': str(pedido_num),
                'TIPO_PEDIDO': str(tipo_pedido),
                'CLIENTE': str(subgrupo['CLIENTE'].iloc[0]),
                'RESERVADO': [],
                'CANCELADO': []
            }
            
            # Separar itens por status
            # RESERVADO
            mask_reservado = (
                ((subgrupo['ULT_STATUS'] == 520) & (subgrupo['PROX_STATUS'] == 540)) |
                ((subgrupo['ULT_STATUS'] == 541) & (subgrupo['PROX_STATUS'] == 540))
            )
            
            # CANCELADO
            mask_cancelado = (
                ((subgrupo['ULT_STATUS'] == 980) & (subgrupo['PROX_STATUS'] == 999)) |
                ((subgrupo['ULT_STATUS'] == 982) & (subgrupo['PROX_STATUS'] == 999)) |
                ((subgrupo['ULT_STATUS'] == 983) & (subgrupo['PROX_STATUS'] == 999))
            )
            
            # Processar RESERVADOS
            itens_reservados = subgrupo[mask_reservado]
            for _, item in itens_reservados.iterrows():
                pedido_novo['RESERVADO'].append({
                    'CODIGO': str(item['ITEM']),
                    'DESCRICAO_CODIGO': str(item['DESC_ITEM']),
                    'UM': str(item['UN_MEDIDA']),
                    'QTDE': str(int(item['QTD'])) if pd.notna(item['QTD']) else '0'
                })
            
            # Processar CANCELADOS
            itens_cancelados = subgrupo[mask_cancelado]
            for _, item in itens_cancelados.iterrows():
                pedido_novo['CANCELADO'].append({
                    'CODIGO': str(item['ITEM']),
                    'DESCRICAO_CODIGO': str(item['DESC_ITEM']),
                    'UM': str(item['UN_MEDIDA']),
                    'QTDE': str(int(item['QTD'])) if pd.notna(item['QTD']) else '0'
                })
            
            resultados[chave]['PEDIDO_NOVO'].append(pedido_novo)
    
    print(f"  ✓ Processados {formatar_numero(len(resultados))} grupos")
    return dict(resultados)

def processar_wms(df_lotes, resultados, mapa_relacionamentos):
    """
    Processa base_lotes e adiciona ao PEDIDO_WMS
    """
    print("\n📦 PROCESSANDO BASE_LOTES (PEDIDO_WMS)...")
    
    if df_lotes.empty:
        print("  ✗ Nenhum dado na base_lotes")
        return resultados
    
    processados = 0
    nao_encontrados = 0
    
    for _, lote in df_lotes.iterrows():
        pedido = lote['pedido']
        tipo = lote['tipo']
        cliente = lote['cod_cliente']
        
        # Buscar NUM_PEDIDO_SISCAP usando o mapa
        chave_busca = (pedido, tipo, cliente)
        
        if chave_busca in mapa_relacionamentos:
            num_pedido_siscap = mapa_relacionamentos[chave_busca]
            chave_resultado = f"{num_pedido_siscap}|{cliente}"
            
            if chave_resultado in resultados:
                # Verificar se já existe este lote
                lote_existente = None
                for wms_item in resultados[chave_resultado]['PEDIDO_WMS']:
                    if wms_item['LOTE'] == str(lote['lote_pedido']) and wms_item['PEDIDO'] == str(pedido):
                        lote_existente = wms_item
                        break
                
                if lote_existente:
                    # Adicionar item ao lote existente
                    lote_existente['ITENS'].append({
                        'CODIGO': str(lote['item']),
                        'DESCRICAO_CODIGO': str(lote['descricao']),
                        'UM': str(lote['um']),
                        'QTDE': str(int(lote['qtde'])) if pd.notna(lote['qtde']) else '0'
                    })
                else:
                    # Criar novo lote
                    resultados[chave_resultado]['PEDIDO_WMS'].append({
                        'LOTE': str(lote['lote_pedido']),
                        'PEDIDO': str(pedido),
                        'TIPO': str(tipo),
                        'ZONA': str(lote['zona']),
                        'ITENS': [{
                            'CODIGO': str(lote['item']),
                            'DESCRICAO_CODIGO': str(lote['descricao']),
                            'UM': str(lote['um']),
                            'QTDE': str(int(lote['qtde'])) if pd.notna(lote['qtde']) else '0'
                        }]
                    })
                
                processados += 1
            else:
                nao_encontrados += 1
        else:
            nao_encontrados += 1
    
    print(f"  ✓ {formatar_numero(processados)} lotes processados")
    if nao_encontrados > 0:
        print(f"  ⚠️  {formatar_numero(nao_encontrados)} registros não relacionados encontrados")
    
    return resultados

def processar_postado(df_postado, resultados, mapa_relacionamentos):
    """
    Processa base_objeto_postado e adiciona ao PEDIDO_POSTADO
    """
    print("\n📮 PROCESSANDO BASE_OBJETO_POSTADO (PEDIDO_POSTADO)...")
    
    if df_postado.empty:
        print("  ✗ Nenhum dado na base_objeto_postado")
        return resultados
    
    processados = 0
    nao_encontrados = 0
    
    for _, objeto in df_postado.iterrows():
        pedido = objeto['PEDIDO']
        tipo = objeto['TIPO_PEDIDO']
        cliente = objeto['CLIENTE']
        
        # Buscar NUM_PEDIDO_SISCAP usando o mapa
        chave_busca = (pedido, tipo, cliente)
        
        if chave_busca in mapa_relacionamentos:
            num_pedido_siscap = mapa_relacionamentos[chave_busca]
            chave_resultado = f"{num_pedido_siscap}|{cliente}"
            
            if chave_resultado in resultados:
                # Adicionar objeto postado
                resultados[chave_resultado]['PEDIDO_POSTADO'].append({
                    'LOTE': str(objeto['LOTE']),
                    'NOTA_FISCAL': str(objeto['NOTA_FISCAL']),
                    'SERIE': str(objeto['SERIE']),
                    'PEDIDO': str(pedido),
                    'TIPO': str(tipo),
                    'ZONA': str(objeto['ZONA']),
                    'OBJETO': str(objeto['REGISTRO']),
                    'DATA_EXPEDICAO': objeto['DATA_EXPEDICAO'].strftime('%Y-%m-%d') if pd.notna(objeto['DATA_EXPEDICAO']) else ''
                })
                
                processados += 1
            else:
                nao_encontrados += 1
        else:
            nao_encontrados += 1
    
    print(f"  ✓ {formatar_numero(processados)} objetos postados processados")
    if nao_encontrados > 0:
        print(f"  ⚠️  {formatar_numero(nao_encontrados)} registros não relacionados encontrados")
    
    return resultados

def salvar_resultados(resultados, caminho_saida):
    """
    Salva resultados em arquivos JSON individuais
    """
    print("\n💾 SALVANDO RESULTADOS...")
    
    os.makedirs(caminho_saida, exist_ok=True)
    
    total_salvo = 0
    
    for chave, estrutura in resultados.items():
        try:
            # Limpar chaves com arrays vazios para JSON mais limpo
            if not estrutura['PEDIDO_WMS']:
                del estrutura['PEDIDO_WMS']
            if not estrutura['PEDIDO_POSTADO']:
                del estrutura['PEDIDO_POSTADO']
            
            # Nome do arquivo seguro
            nome_safe = chave.replace('|', '_').replace('\\', '_').replace('/', '_')
            nome_arquivo = f"pedido_{nome_safe}.json"
            caminho_completo = os.path.join(caminho_saida, nome_arquivo)
            
            # Salvar JSON
            with open(caminho_completo, 'w', encoding='utf-8') as f:
                json.dump(estrutura, f, ensure_ascii=False, indent=2)
            
            total_salvo += 1
            
        except Exception as e:
            print(f"  ✗ Erro ao salvar {chave}: {e}")
    
    # Salvar resumo em Excel
    salvar_resumo_excel(resultados, caminho_saida)
    
    print(f"  ✓ {formatar_numero(total_salvo)} arquivos JSON salvos em: {caminho_saida}")

def salvar_resumo_excel(resultados, caminho_saida):
    """
    Cria um resumo em Excel com estatísticas
    """
    dados_resumo = []
    
    for chave, estrutura in resultados.items():
        for pedido in estrutura.get('PEDIDO_NOVO', []):
            dados_resumo.append({
                'NUM_PEDIDO_SISCAP': estrutura['NUM_PEDIDO_SISCAP'],
                'CLIENTE': estrutura['CLIENTE'],
                'NOME_CLIENTE': estrutura['NOME_CLIENTE'],
                'MCCO': estrutura['MCCO'],
                'PEDIDO': pedido['PEDIDO'],
                'TIPO_PEDIDO': pedido['TIPO_PEDIDO'],
                'DATA_PEDIDO': pedido['DATA_PEDIDO'],
                'QTD_RESERVADOS': len(pedido['RESERVADO']),
                'QTD_CANCELADOS': len(pedido['CANCELADO']),
                'QTD_LOTES_WMS': len(estrutura.get('PEDIDO_WMS', [])),
                'QTD_OBJETOS_POSTADOS': len(estrutura.get('PEDIDO_POSTADO', []))
            })
    
    if dados_resumo:
        df_resumo = pd.DataFrame(dados_resumo)
        caminho_excel = os.path.join(caminho_saida, '00_RESUMO_PROCESSAMENTO.xlsx')
        
        # Criar Excel writer
        with pd.ExcelWriter(caminho_excel, engine='openpyxl') as writer:
            df_resumo.to_excel(writer, sheet_name='Resumo', index=False)
            
            # Adicionar estatísticas
            total_grupos = len(resultados)
            total_pedidos = len(dados_resumo)
            total_reservados = df_resumo['QTD_RESERVADOS'].sum()
            total_cancelados = df_resumo['QTD_CANCELADOS'].sum()
            total_lotes = df_resumo['QTD_LOTES_WMS'].sum()
            total_postados = df_resumo['QTD_OBJETOS_POSTADOS'].sum()
            
            stats = {
                'Métrica': [
                    'Total de grupos processados',
                    'Total de pedidos únicos',
                    'Total de itens reservados',
                    'Total de itens cancelados',
                    'Total de lotes WMS',
                    'Total de objetos postados'
                ],
                'Valor': [
                    total_grupos,
                    total_pedidos,
                    total_reservados,
                    total_cancelados,
                    total_lotes,
                    total_postados
                ]
            }
            
            df_stats = pd.DataFrame(stats)
            df_stats.to_excel(writer, sheet_name='Estatísticas', index=False)
        
        print(f"  ✓ Resumo salvo em Excel: {caminho_excel}")

def testar_pequeno_conjunto(caminho_excel):
    """
    Testa o processamento com um conjunto pequeno de dados
    """
    print("\n🔬 TESTE COM 100 REGISTROS")
    print("="*60)
    
    # Ler apenas 100 registros para teste
    planilhas = ler_todas_planilhas(caminho_excel, limite=100)
    
    if planilhas:
        # Criar mapa de relacionamentos
        mapa = criar_mapa_relacionamentos(planilhas['pedidos'])
        
        # Processar pedidos
        resultados = processar_pedidos_novos(planilhas['pedidos'], mapa)
        
        # Processar WMS
        resultados = processar_wms(planilhas['lotes'], resultados, mapa)
        
        # Processar postado
        resultados = processar_postado(planilhas['postado'], resultados, mapa)
        
        if resultados:
            # Mostrar exemplo
            primeira_chave = list(resultados.keys())[0]
            print(f"\n📄 EXEMPLO DO PRIMEIRO JSON:")
            exemplo = json.dumps(resultados[primeira_chave], indent=2, ensure_ascii=False)
            print(exemplo[:800] + "..." if len(exemplo) > 800 else exemplo)
            
            # Mostrar estrutura
            print(f"\n📊 ESTRUTURA DO PRIMEIRO REGISTRO:")
            estrutura = resultados[primeira_chave]
            print(f"  • NUM_PEDIDO_SISCAP: {estrutura['NUM_PEDIDO_SISCAP']}")
            print(f"  • CLIENTE: {estrutura['CLIENTE']}")
            print(f"  • PEDIDO_NOVO: {len(estrutura['PEDIDO_NOVO'])} pedido(s)")
            print(f"  • PEDIDO_WMS: {len(estrutura.get('PEDIDO_WMS', []))} lote(s)")
            print(f"  • PEDIDO_POSTADO: {len(estrutura.get('PEDIDO_POSTADO', []))} objeto(s)")
            
            return True
    
    return False

def main():
    """
    Função principal
    """
    print("="*80)
    print("PROCESSADOR COMPLETO DE EXCEL PARA JSON - VERSÃO 2.0")
    print("Processa: base_pedidos + base_lotes + base_objeto_postado")
    print("="*80)
    
    # Configurar
    configurar_pandas()
    
    # Caminhos
    caminho_excel = r"c:\app\dashboard\001-base_completa.xlsm"
    caminho_saida = "resultados_json_completo"
    
    # Verificar arquivo
    if not os.path.exists(caminho_excel):
        print(f"❌ Arquivo não encontrado: {caminho_excel}")
        print("Por favor, ajuste o caminho do arquivo.")
        return
    
    print(f"📂 Arquivo de entrada: {caminho_excel}")
    print(f"📁 Pasta de saída: {caminho_saida}")
    
    # Teste inicial
    print("\n🎯 FASE 1: TESTE INICIAL")
    if not testar_pequeno_conjunto(caminho_excel):
        print("❌ Teste inicial falhou. Verifique o arquivo Excel.")
        return
    
    # Opções de processamento
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
            limite = 10000
        elif opcao == '3':
            limite = None
        elif opcao == '4':
            limite = int(input("Digite a quantidade: "))
        else:
            print("Opção inválida. Usando padrão (1.000 registros).")
            limite = 1000
    except:
        print("Entrada inválida. Usando padrão (1.000 registros).")
        limite = 1000
    
    # Processamento principal
    if limite is None:
        print(f"\n🚀 FASE 2: PROCESSAMENTO PRINCIPAL (TODOS os registros)")
    else:
        print(f"\n🚀 FASE 2: PROCESSAMENTO PRINCIPAL ({formatar_numero(limite)} registros)")
    
    inicio_total = time.time()
    
    # 1. Ler todas as planilhas
    planilhas = ler_todas_planilhas(caminho_excel, limite)
    
    if not planilhas:
        print("❌ Falha ao ler planilhas.")
        return
    
    # 2. Criar mapa de relacionamentos
    mapa = criar_mapa_relacionamentos(planilhas['pedidos'])
    
    # 3. Processar pedidos novos
    resultados = processar_pedidos_novos(planilhas['pedidos'], mapa)
    
    # 4. Processar base_lotes (PEDIDO_WMS)
    resultados = processar_wms(planilhas['lotes'], resultados, mapa)
    
    # 5. Processar base_objeto_postado (PEDIDO_POSTADO)
    resultados = processar_postado(planilhas['postado'], resultados, mapa)
    
    # 6. Salvar resultados
    if resultados:
        salvar_resultados(resultados, caminho_saida)
    else:
        print("❌ Nenhum resultado para salvar.")
        return
    
    # Estatísticas finais
    fim_total = time.time()
    tempo_total = fim_total - inicio_total
    
    print("\n" + "="*80)
    print("RESUMO FINAL DO PROCESSAMENTO")
    print("="*80)
    
    # Calcular estatísticas
    total_pedidos = sum(len(g['PEDIDO_NOVO']) for g in resultados.values())
    total_itens_reservados = sum(
        len(p['RESERVADO'])
        for g in resultados.values()
        for p in g['PEDIDO_NOVO']
    )
    total_itens_cancelados = sum(
        len(p['CANCELADO'])
        for g in resultados.values()
        for p in g['PEDIDO_NOVO']
    )
    total_lotes = sum(len(g.get('PEDIDO_WMS', [])) for g in resultados.values())
    total_postados = sum(len(g.get('PEDIDO_POSTADO', [])) for g in resultados.values())
    
    print(f"📊 Total de grupos (NUM_PEDIDO_SISCAP+CLIENTE): {formatar_numero(len(resultados))}")
    print(f"📊 Total de pedidos únicos: {formatar_numero(total_pedidos)}")
    print(f"📊 Total de itens RESERVADOS: {formatar_numero(total_itens_reservados)}")
    print(f"📊 Total de itens CANCELADOS: {formatar_numero(total_itens_cancelados)}")
    print(f"📊 Total de lotes WMS: {formatar_numero(total_lotes)}")
    print(f"📊 Total de objetos postados: {formatar_numero(total_postados)}")
    print(f"⏱️  Tempo total de processamento: {tempo_total:.2f} segundos")
    
    if limite:
        velocidade = limite / tempo_total if tempo_total > 0 else 0
        print(f"⚡ Velocidade: {velocidade:.0f} registros/segundo")
    
    print(f"\n✅ PROCESSAMENTO CONCLUÍDO COM SUCESSO!")
    print(f"📁 Resultados salvos em: {os.path.abspath(caminho_saida)}")
    print("="*80)
    
    # Mostrar alguns exemplos
    print("\n📋 EXEMPLOS DE ARQUIVOS GERADOS:")
    try:
        arquivos = [f for f in os.listdir(caminho_saida) if f.endswith('.json')]
        for i, arquivo in enumerate(arquivos[:5], 1):
            print(f"  {i}. {arquivo}")
        if len(arquivos) > 5:
            print(f"  ... e mais {len(arquivos) - 5} arquivos")
    except:
        pass

if __name__ == "__main__":
    main()