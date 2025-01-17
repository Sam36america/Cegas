import os
import PyPDF2
import re
import pandas as pd
from openpyxl import load_workbook
import shutil
import xml.etree.ElementTree as ET

DIST = 'Cegás'

class ExtratorFaturas:
    def __init__(self):
        self.regexes = {
            'cnpj': [r'(\d{2}\.\d{3}\.?\d{3}\/?\d{4}\-?\s?\d{2})\s+\d{2}\/\d{2}\/\d{4}'], #26/09
            'valor_total': [r'-?(\d+\.?\d+\,\d{2})\s\d{2}\/\d{2}\/\d{4}'], #26/09
            'volume_total': [r'M3\s(\d+\.?\,?\d+\.?\,?\d+)\s?'], #26/09
            'data_emissao': [r'[A]\s(\d{2}\/\d{2}\/\d{4})'], #05/11
            'data_inicio': [r'DE\s(\d{2}\/\d{2}\/\d{4})\s'], #26/09
            'data_fim': [r'[A-a]\s(\d{2}\/\d{2}\/\d{4})'], #26/09            
            'numero_fatura': [r'Nº\s(\d+\.?\d+\.?\d+)\s'], #26/09
            'valor_icms': [r'\,\d+\s(\d+\d+\,\.?\d+)\s[0]'], #05/12    
            'correcao_pcs': [r'R\d+\s(\d+)\s*', r'T\d+\s(\d+)\s*'] # DIVIDIR ESSA MERDA POR 9400 E PEGAR 4 CASAS DECIMAIS APÓS O 0.
        }

    def extrair_informacoes(self, texto):
        informacoes = {}
        for chave, regex_list in self.regexes.items():
            for regex in regex_list:
                match = re.search(regex, texto)
                if match:
                    valor = match.group(1) if match.groups() else match.group(0)
                    if chave == 'correcao_pcs':
                        try:
                            valor = float(valor) / 9400
                            valor = f"{valor:.4f}"
                        except ValueError:
                            valor = ''
                    informacoes[chave] = valor
                    break  # Para de procurar assim que encontrar uma correspondência
        return informacoes

def extrair_texto(caminho_do_pdf):
    texto = ''
    try:
        with open(caminho_do_pdf, 'rb') as arquivo:
            leitor_pdf = PyPDF2.PdfReader(arquivo)
            for pagina in leitor_pdf.pages:
                texto_pagina = pagina.extract_text()
                if texto_pagina:
                    texto_pagina = texto_pagina.replace('\n', ' ')
                    texto_pagina = re.sub(r'\s{2,}', ' ', texto_pagina).strip()
                    texto += texto_pagina + ' '
    except Exception as e:
        print(f"Erro ao ler o PDF: {e}")
    
    if not texto:
        print(f"Erro ao extrair texto do PDF: {caminho_do_pdf}")
    else:
        print(f"Texto extraído com sucesso...")  # Mostra os primeiros 900 caracteres do texto extraído
    return texto.strip()  # Remove espaços extras no início e no fim

def extrair_texto_xml(caminho_do_xml):
    texto = ''
    try:
        tree = ET.parse(caminho_do_xml)
        root = tree.getroot()
        texto = ET.tostring(root, encoding='unicode', method='text')
    except Exception as e:
        print(f"Erro ao ler o XML: {e}")
    
    if not texto:
        print(f"Erro ao extrair texto do XML: {caminho_do_xml}")
    else:
        print(f"Texto extraído com sucesso do XML...")
    return texto.strip()

def registro_existe(df, cnpj, data_inicio, data_fim, valor_total):
    return not df[(df['CNPJ'] == cnpj) & (df['Data Início'] == data_inicio) & (df['Data Fim'] == data_fim) & (df['Valor Total'] == valor_total)].empty

def adicionar_na_planilha(informacoes, caminho_planilha, nome_arquivo):
    try:
        df = pd.read_excel(caminho_planilha)
    except FileNotFoundError:
        print(f"O arquivo '{caminho_planilha}' não foi encontrado. Criando um novo.")
        df = pd.DataFrame(columns=['CNPJ', 'Valor Total', 'Volume Total', 'Data Emissão', 'Data Início', 'Data Fim', 'Número Fatura', 'Valor ICMS', 'Correção PCS', 'Distribuidora', 'Nome do Arquivo'])
    
    cnpj = informacoes['cnpj']
    data_inicio = informacoes['data_inicio']
    data_fim = informacoes['data_fim']
    valor_total = pd.to_numeric(informacoes['valor_total'].replace('.', '').replace(',', '.'))
    volume_total = pd.to_numeric(informacoes['volume_total'].replace('.', '').replace(',', '.'))
    valor_icms = pd.to_numeric(informacoes['valor_icms'].replace('.', '').replace(',', '.'))

    if registro_existe(df, cnpj, data_inicio, data_fim, valor_total):
        print(f"Registro duplicado encontrado para CNPJ: {cnpj}, Data Início: {data_inicio}, Data Fim: {data_fim}, Valor Total: {valor_total}. Não será inserido.")
        return False 

    # Converte os valores para numéricos
    valor_total = pd.to_numeric(str(informacoes.get('valor_total', '')).replace('.', '').replace(',', '.'), errors='coerce')
    volume_total = pd.to_numeric(str(informacoes.get('volume_total', '')).replace('.', '').replace(',', '.'), errors='coerce')
    valor_icms = pd.to_numeric(str(informacoes.get('valor_icms', '')).replace('.', '').replace(',', '.'), errors='coerce')
    correcao_pcs = pd.to_numeric(informacoes.get('correcao_pcs', ''), errors='coerce')
    
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
        'Nome do Arquivo': nome_arquivo  # Adiciona o nome do arquivo
    }])
    df = pd.concat([df, nova_linha], ignore_index=True)
    df.to_excel(caminho_planilha, index=False)
    return True  # Indica que o registro foi inserido

def mover_arquivo(origem, destino):
    shutil.move(origem, destino)
    print(f"Arquivo movido para {destino}")

def main(file_path, file, caminho_planilha):
    if file.lower().endswith('.pdf'):
        texto = extrair_texto(file)
    elif file.lower().endswith('.xml'):
        texto = extrair_texto_xml(file)
    else:
        print(f"Tipo de arquivo não suportado: {file}")
        return

    if not texto:
        print(f"Erro ao extrair texto do arquivo: {file}")
        return

    extrator = ExtratorFaturas()
    informacoes = extrator.extrair_informacoes(texto)
    
    # Verifica se todos os campos foram extraídos
    campos_necessarios = ['cnpj', 'valor_total', 'volume_total', 'data_emissao', 'data_inicio', 'data_fim', 'numero_fatura', 'valor_icms', 'correcao_pcs']
    campos_faltantes = [campo for campo in campos_necessarios if not informacoes.get(campo)]
    
    if campos_faltantes:
        print(f"Campos faltantes no arquivo {file}: {', '.join(campos_faltantes)}")
        return

    nome_arquivo = os.path.basename(file)  # Extrai apenas o nome do arquivo
    inserido = adicionar_na_planilha(informacoes, caminho_planilha, nome_arquivo)
    print(informacoes)

    if inserido:
        destino = os.path.join(diretorio_destino, nome_arquivo)
        mover_arquivo(file, destino)
    else:
        print('Arquivo já foi inserido na planilha. Não será movido.')
    
# Exemplo de uso
file_path = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\Cegás\Faturas'
diretorio_destino = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\Cegás\Lidos'
caminho_planilha = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\00 Faturas Lidas\CEGAS.xlsx'

for arquivo in os.listdir(file_path):
    if arquivo.lower().endswith('.pdf') or arquivo.lower().endswith('.xml'):
        arquivo_full = os.path.join(file_path, arquivo)
        arquivo = os.path.basename(arquivo)

        main(file_path, arquivo_full, caminho_planilha)