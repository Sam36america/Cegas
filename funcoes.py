from pdf2image import convert_from_path
import pandas as pd
import numpy as np
import os
import re
from sqlalchemy import create_engine, select, MetaData, Table
from sqlalchemy.orm import sessionmaker
import shutil
import PyPDF2

#in/out 3100, 1300, 3800, 1450
usuario_conectado = 'samuel.santos'
# Configure o caminho do executável do Tesseract

      
def dados_excel(cnpj, valor_total,volume_total, data_emissao, data_inicio, data_fim, numero_fatura, valor_icms, correcao_pcs, dist):
   
    dados = {
           'CNPJ': [cnpj],
           'VALOR TOTAL': valor_total,
           'VOLUME TOTAL': [volume_total],
           'DATA DA EMISSÃO': data_emissao,
           'DATA INICIO': [data_inicio],
           'DATA FIM': [data_fim],     
           'NUMERO FATURA':[numero_fatura],
           'VALOR ICMS': [valor_icms],
           'CORREÇÃO PCS': [correcao_pcs],
           'DISTRIBUIDORA': [dist]
     }
    try:    
        df = pd.DataFrame(dados)
    except:
        dados = {
           'CNPJ':'CNPJ não encontrado', 
           'VALOR TOTAL': 'valor_total não econtrado',
           'VOLUME TOTAL': 'volume_total não econtrado',
           'DATA DA EMISSÃO': 'data_emissao não econtrado',
           'DATA INICIO': 'data_inicio não econtrado',
           'DATA FIM': 'data_fim não econtrado',     
           'NUMERO FATURA':'numero_fatura não econtrado]',
           'VALOR ICMS': 'valor_icms não econtrado',
           'CORREÇÃO DO PCS': 'correcao_pcs não encontrado',
           'DISTRIBUIDORA': 'Distribuidora não econtrado'
           
            }
        
        indice = ['1']
        df = pd.DataFrame(dados,index=indice)  
    
    return df
          
def adicionar_dados_excel(dados, novos_dados):
    try:
        df_existente = pd.read_excel(dados)
        
    except FileNotFoundError:
        print(f"O arquivo '{dados}' não foi encontrado. Criando um novo.")
        df_existente = pd.DataFrame()
    
    try:
        df_novos_dados = pd.DataFrame(novos_dados)
        df_resultante = pd.concat([df_existente, df_novos_dados], ignore_index=True)
        df_resultante.to_excel(dados, index=False)
        print(f"Dados adicionados com sucesso na planilha '{dados}'")
        return True
    except:
        print(f"Erro ao adicionados os dados na planilha '{dados}'")
        return False

def listar_pdfs_com_referencia_na_pasta(pasta, referencia):
    arquivos_pdf = []
    for arquivo in os.listdir(pasta):
        if arquivo.endswith('.pdf'):
            nome_distribuidora = re.findall(r'_GN_([A-ZÁ]+)_',arquivo)
            if nome_distribuidora:
                nome_distribuidora = nome_distribuidora[0]
                
            arquivos_pdf.append(arquivo)
    return arquivos_pdf

def verificar_fatura_existe(session, tabela_faturas, numero_fatura):
    stmt = select([tabela_faturas.c.numero_fatura]).where(tabela_faturas.c.numero_fatura == numero_fatura)
    result = session.execute(stmt).fetchone()
    return result is not None

def verificar_download(cnpj, data_inicio, data_fim, excel_path):
    # Carregar o arquivo Excel
    df = pd.read_excel(excel_path, sheet_name='Sheet1')
    
    cnpj = int(cnpj)

    # Filtrar as linhas que correspondem aos critérios
    df_filtrado = df[
        (df['CNPJ'] == cnpj) &
        (df['DATA INICIO'] == data_inicio) &
        (df['DATA FIM'] == data_fim)
    ]
    
    class ExtratorFaturas:
        def __init__(self):
            self.regexes = {
                'cnpj': r'\d{2}\.\d{3}\.?\d{3}\/?\d{4}\-?\s?\d{2}',
                'valor_total': r'R\$\s(\d+\.?\d+\,\d{2})\s',
                'volume_total': r'Total\s\d+\.?\,?\d+\.?\,?\d+?\.?\,?\s',
                'data_emissao': r'apresentação\s(\d{2}\.\d{2}\.\d{4})',
                'data_inicio': r'\d{2}\.\d{2}\.\d{4}\d{2}\.\d{2}\.\d{4}(\d{2}\.\d{2}\.\d{4})\d{2}\.\d{2}\.\d{4}', #revisar essa merda de regex
                'data_fim': r'\d{2}\.\d{2}\.\d{4}(\d{2}\.\d{2}\.\d{4})\d{2}\.\d{2}\.\d{4}',            
                'numero_fatura': r'\s(\d{3}\.\d{3}\.\d{3})\s',
                'valor_icms': r'ICMS\s?R\$\s(\d+\.?\d+\,\d{2})\s',     
                'correcao_pcs': r'[A-Z]\d{9}(\d{4})\d+',
            }
    
        def extrair_informacoes(self, texto):
            informacoes = {}
            for chave, regex in self.regexes.items():
                match = re.search(regex, texto)
                if match:
                    informacoes[chave] = match.group(1) if match.groups() else match.group(0)
                else:
                    informacoes[chave] = ''  # Adiciona uma string vazia se não houver correspondência
            return informacoes
    
    def extrair_texto(caminho_do_pdf):
        texto = ''
        with open(caminho_do_pdf, 'rb') as arquivo:
            leitor_pdf = PyPDF2.PdfReader(arquivo)
            for pagina in leitor_pdf.pages:
                texto_pagina = pagina.extract_text()
                if texto_pagina:
                    texto += texto_pagina + ' '  # Substitui '\n' por ' ' para evitar linhas extras
        return texto.strip()  # Remove espaços extras no início e no fim
    
    def adicionar_na_planilha(informacoes, caminho_planilha, nome_arquivo):
        df = pd.read_excel(caminho_planilha)
        nova_linha = pd.DataFrame([{
            'CNPJ': informacoes.get('cnpj', ''),
            'Valor Total': informacoes.get('valor_total', ''),
            'Volume Total': informacoes.get('volume_total', ''),
            'Data Emissão': informacoes.get('data_emissao', ''),
            'Data Início': informacoes.get('data_inicio', ''),
            'Data Fim': informacoes.get('data_fim', ''),
            'Número Fatura': informacoes.get('numero_fatura', ''),
            'Valor ICMS': informacoes.get('valor_icms', ''),
            'Correção PCS': informacoes.get('correcao_pcs', ''),
            'Nome Arquivo': nome_arquivo  # Adiciona o nome do arquivo
        }])
        df = pd.concat([df, nova_linha], ignore_index=True)
        df.to_excel(caminho_planilha, index=False)
    
    def main(file_path, pdf_file, caminho_planilha):
        texto_pypdf = extrair_texto(pdf_file)
        if not texto_pypdf:
            print(f"Erro ao extrair texto do PDF: {pdf_file}")
            return
    
        extrator = ExtratorFaturas()
        informacoes = extrator.extrair_informacoes(texto_pypdf)
        if not any(informacoes.values()):
            print(f"Nenhuma informação extraída do PDF: {pdf_file}")
            return
    
        adicionar_na_planilha(informacoes, caminho_planilha, os.path.basename(pdf_file))
        #print(informacoes)
    
    # Exemplo de uso
    file_path = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\Cegás\Faturas'
    diretorio_destino = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\Cegás\Lidos'
    caminho_planilha = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\Cegás\informacoes_faturas.xlsx'
    
    for arquivo in os.listdir(file_path):
        if arquivo.endswith('.pdf') or arquivo.endswith('.PDF'):
            arquivo_full = rf'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\Cegás\Faturas\{arquivo}' 
            main(arquivo, arquivo_full, caminho_planilha)



