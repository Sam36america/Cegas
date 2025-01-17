# COORDENADAS PARA EXTRAÇÃO DE DADOS USANDO OCR
                    # A PRIMEIRA COORDENADA DE CADA VALOR, FOI FEITA ATRAVÉZ DA FATURA PADRÃO
def corte_comgas():
    corte = {
                                                 
    'cnpj': (2240, 1400, 3200, 1570),                      # PADRÃO                                          
    'cnpj_ajustado': (1700, 1500, 3500, 1620),
    'cnpj_ajustado2': (),

    'valor_total0': (800, 4350, 1380, 4550),
    'valor_total': (3360, 2200, 3820, 2390),               # PADRÃO
    'valor_total_ajustado':(3400, 2240, 3790, 2350),
    'valor_total_ajustado2':(),

    'volume_total': (1980, 3020, 2245, 3150),              # PADRÃO
    'volume_total_ajustado':  (2040, 2990, 2325, 3100),
    'volume_total_ajustado2': (2050, 2910, 2350, 3100),
    'volume_total_ajustado3': (2190, 3030, 2370, 3100),
     
    'data_emissao': (3100, 1395, 3800, 1550),              # PADRÃO
    'data_emissao2': (3200, 1450, 3800, 1600),
    'data_emissao3': (3200, 1550, 3800, 1800),              

    'data_inicio': (85, 4780, 1400, 4900),                 # PADRÃO
    'data_inicio_ajustado': (505, 4680, 1800, 4830),
    'data_inicio_ajustado2': (100, 4650, 1700, 4750),
    'data_inicio_ajustado3': (100, 3000, 1700, 4600),
    
    'data_fim': (85, 4780, 1600, 4900),                    # PADRÃO
    'data_fim_ajustado': (272, 4680, 1800, 4830),
    'data_fim_ajustado2': (300, 4650, 1700, 4750),
    'data_fim_ajustado3': ( 1200, 4620, 1450, 4750),
    'data_fim_ajustado4': (100, 3400, 1700, 4400),
    
    'numero_fatura': (2400, 200, 3750, 450),               # PADRÃO
    'numero_fatura_ajustado': (000, 430, 3750, 590),
    'numero_fatura_ajustado2': (200, 300, 3750,570),
    'numero_fatura_ajustado3': (3000, 425, 3750, 495)  
    }
    return corte

# CAMINHO DA PLANILHA

caminho_excel = r'G:\QUALIDADE\Códigos\Leitura de Faturas Gás\Códigos\dados.xlsx'

'''while True
    for Player(get_moeda):
        if player(moeda >= 3):
            brilhar(dourado)
        else:
            brilhar(branco)'''


