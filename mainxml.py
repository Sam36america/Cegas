import os
import shutil
import xml.etree.ElementTree as ET
import pandas as pd
import re

DIST = 'Cegás'
NAMESPACE = {'nfe': 'http://www.portalfiscal.inf.br/nfe'}

def extrair_informacoes_xml(caminho_do_xml):
    informacoes = {}
    try:
        tree = ET.parse(caminho_do_xml)
        root = tree.getroot()
        
        # Extrair informações básicas do XML
        informacoes['cnpj'] = root.find('.//nfe:emit/nfe:CNPJ', NAMESPACE).text
        informacoes['valor_total'] = root.find('.//nfe:total/nfe:ICMSTot/nfe:vNF', NAMESPACE).text
        informacoes['volume_total'] = root.find('.//nfe:det/nfe:prod/nfe:qCom', NAMESPACE).text
        informacoes['data_emissao'] = root.find('.//nfe:ide/nfe:dhEmi', NAMESPACE).text.split('T')[0]
        inf_cpl = root.find('.//nfe:infAdic/nfe:infCpl', NAMESPACE).text
        informacoes['data_inicio'] = inf_cpl.split(' ')[2]
        informacoes['data_fim'] = inf_cpl.split(' ')[4]
        informacoes['numero_fatura'] = root.find('.//nfe:ide/nfe:nNF', NAMESPACE).text
        informacoes['valor_icms'] = root.find('.//nfe:total/nfe:ICMSTot/nfe:vICMS', NAMESPACE).text

        # Buscar PCS em várias localizações possíveis
        pcs = None
        
        # 1. Tentar encontrar na seção de combustíveis
        comb = root.find('.//nfe:det/nfe:prod/nfe:comb', NAMESPACE)
        if comb is not None:
            for elem in comb.iter():
                if 'PCS' in elem.tag or (elem.text and 'PCS' in elem.text):
                    pcs = elem.text
                    break
        
        # 2. Procurar nas informações adicionais do produto
        if not pcs:
            det = root.find('.//nfe:det', NAMESPACE)
            if det is not None:
                inf_ad_prod = det.find('.//nfe:infAdProd', NAMESPACE)
                if inf_ad_prod is not None and inf_ad_prod.text:
                    pcs_match = re.search(r'PCS[:\s]+(\d+(?:[.,]\d+)?)', inf_ad_prod.text)
                    if pcs_match:
                        pcs = pcs_match.group(1)

        # 3. Procurar no infCpl
        if not pcs and inf_cpl:
            pcs_match = re.search(r'PCS[:\s]+(\d+(?:[.,]\d+)?)', inf_cpl)
            if pcs_match:
                pcs = pcs_match.group(1)

        # 4. Procurar em qualquer lugar do XML que contenha "PCS"
        if not pcs:
            for elem in root.iter():
                if elem.text and 'PCS' in elem.text:
                    pcs_match = re.search(r'(\d+(?:[.,]\d+)?)', elem.text)
                    if pcs_match:
                        pcs = pcs_match.group(1)
                        break

        informacoes['correcao_pcs'] = pcs if pcs else ''

        # Debug: imprimir todos os elementos para análise
        #print("\nDebug - Todos os elementos do XML:")
        for elem in root.iter():
            if elem.text and elem.text.strip():
                print(f"{elem.tag.split('}')[-1]}: {elem.text}")

    except Exception as e:
        print(f"Erro ao extrair informações do XML: {e}")
    
    return informacoes

def registro_existe(df, cnpj, data_inicio, data_fim, valor_total):
    return not df[(df['CNPJ'] == cnpj) & 
                 (df['Data Início'] == data_inicio) & 
                 (df['Data Fim'] == data_fim) & 
                 (df['Valor Total'] == valor_total)].empty

def adicionar_na_planilha(informacoes, caminho_planilha, nome_arquivo):
    try:
        df = pd.read_excel(caminho_planilha)
    except FileNotFoundError:
        print(f"O arquivo '{caminho_planilha}' não foi encontrado. Criando um novo.")
        df = pd.DataFrame(columns=['CNPJ', 'Valor Total', 'Volume Total', 'Data Emissão', 
                                 'Data Início', 'Data Fim', 'Número Fatura', 'Valor ICMS', 
                                 'Correção PCS', 'Distribuidora', 'Nome do Arquivo'])
    
    cnpj = informacoes['cnpj']
    data_inicio = informacoes['data_inicio']
    data_fim = informacoes['data_fim']
    valor_total = pd.to_numeric(informacoes['valor_total'].replace('.', '').replace(',', '.'))
    volume_total = pd.to_numeric(informacoes['volume_total'].replace('.', '').replace(',', '.'))
    valor_icms = pd.to_numeric(informacoes['valor_icms'].replace('.', '').replace(',', '.'))

    if registro_existe(df, cnpj, data_inicio, data_fim, valor_total):
        print(f"Registro duplicado encontrado para CNPJ: {cnpj}, Data Início: {data_inicio}, "
              f"Data Fim: {data_fim}, Valor Total: {valor_total}. Não será inserido.")
        return False 

    # Extrair e dividir correção PCS
    try:
        correcao_pcs = float(informacoes.get('correcao_pcs', '').replace('.', '').replace(',', '.')) / 9400
        correcao_pcs = f"{correcao_pcs:.4f}"
    except ValueError:
        correcao_pcs = ''

    nova_linha = pd.DataFrame([{
        'CNPJ': cnpj,
        'Valor Total': valor_total,
        'Volume Total': volume_total,
        'Data Emissão': informacoes.get('data_emissao', ''),
        'Data Início': data_inicio,
        'Data Fim': data_fim,
        'Número Fatura': informacoes.get('numero_fatura', ''),
        'Valor ICMS': valor_icms,
        'Correção PCS': correcao_pcs,
        'Distribuidora': DIST,
        'Nome do Arquivo': nome_arquivo
    }])
    
    df = pd.concat([df, nova_linha], ignore_index=True)
    df.to_excel(caminho_planilha, index=False)
    return True

def mover_arquivo(origem, destino):
    shutil.move(origem, destino)
    print(f"Arquivo movido para {destino}")

def processar_xml(file, caminho_planilha, diretorio_destino):
    informacoes = extrair_informacoes_xml(file)
    
    # Verifica se todos os campos foram extraídos
    campos_necessarios = ['cnpj', 'valor_total', 'volume_total', 'data_emissao', 
                         'data_inicio', 'data_fim', 'numero_fatura', 'valor_icms', 'correcao_pcs']
    campos_faltantes = [campo for campo in campos_necessarios if not informacoes.get(campo)]
    
    if campos_faltantes:
        print(f"Campos faltantes no arquivo {file}: {', '.join(campos_faltantes)}")
        return

    nome_arquivo = os.path.basename(file)
    inserido = adicionar_na_planilha(informacoes, caminho_planilha, nome_arquivo)
    print(informacoes)

    if inserido:
        destino = os.path.join(diretorio_destino, nome_arquivo)
        mover_arquivo(file, destino)
    else:
        print('Arquivo já foi inserido na planilha. Não será movido.')

# Caminhos dos diretórios
file_path = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\Cegás\Faturas'
diretorio_destino = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\Cegás\Lidos'
caminho_planilha = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\00 Faturas Lidas\CEGAS.xlsx'

# Processamento principal
for arquivo in os.listdir(file_path):
    if arquivo.lower().endswith('.xml'):
        arquivo_full = os.path.join(file_path, arquivo)
        processar_xml(arquivo_full, caminho_planilha, diretorio_destino)