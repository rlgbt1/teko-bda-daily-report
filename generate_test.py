from src.report_generator.pptx_builder import BDAReportGenerator

data = {
    'report_date': '01.04.2026',
    'reembolso_credito': '17,62 M Kz',
    'kpis': [
        {'label': 'Liquidez MN',          'value': '45,3 mM Kz',  'variation_str': '+2,1%'},
        {'label': 'Liquidez ME',           'value': '12,5 M USD',  'variation_str': '-0,5%'},
        {'label': 'Posição Cambial',       'value': '5.230 mM Kz', 'variation_str': '+1,2%'},
        {'label': 'Carteira Títulos',      'value': '320 mM Kz',   'variation_str': ''},
        {'label': 'Rentabilidade MN',      'value': '8,75%',       'variation_str': '+0,1%'},
        {'label': 'Rentabilidade ME',      'value': '3,20%',       'variation_str': ''},
        {'label': 'Rentabilidade Títulos', 'value': '9,10%',       'variation_str': ''},
        {'label': 'Reembolsos',            'value': '17,62 M Kz',  'variation_str': ''},
    ],
    'liquidez_mn_days': ['26/03', '27/03', '28/03', '31/03', '01/04'],
    'liquidez_mn_rows': [
        {'label': 'Posição Reservas Livres BNA', 'values': ['12.500', '13.200', '12.800', '13.100', '13.450']},
        {'label': 'Posição DO B. Comerciais',    'values': ['8.300',  '8.100',  '8.500',  '8.200',  '8.400']},
        {'label': 'Posição DP B. Comerciais',    'values': ['5.200',  '5.300',  '5.100',  '5.400',  '5.350']},
        {'label': 'Posição OMAs',                'values': ['2.100',  '2.200',  '2.050',  '2.150',  '2.200']},
        {'label': 'LIQUIDEZ BDA',                'values': ['28.100', '28.800', '28.450', '28.850', '29.400']},
    ],
    'transacoes_mn_raw': [
        {'tipo': 'OMA', 'contraparte': 'BNA', 'taxa': '19,50%', 'montante': '2.200 mM Kz', 'maturidade': '02/04/2026', 'juros': '115 mM Kz'},
    ],
    'operacoes_vivas': [
        {'tipo': 'DP', 'contraparte': 'BAI', 'montante': '5.000 mM Kz', 'taxa': '20,00%', 'residual': 15, 'vencimento': '16/04/2026', 'juro_diario': '27,4 mM Kz'},
        {'tipo': 'DP', 'contraparte': 'BFA', 'montante': '3.500 mM Kz', 'taxa': '19,75%', 'residual': 30, 'vencimento': '01/05/2026', 'juro_diario': '18,9 mM Kz'},
    ],
    'luibor': {
        'LUIBOR O/N':  '19,25%',
        'LUIBOR 1M':   '19,50%',
        'LUIBOR 3M':   '19,75%',
        'LUIBOR 6M':   '20,00%',
        'LUIBOR 9M':   '20,25%',
        'LUIBOR 12M':  '20,50%',
    },
    'luibor_d1': {
        'LUIBOR O/N':  '19,25%',
        'LUIBOR 1M':   '19,45%',
        'LUIBOR 3M':   '19,75%',
        'LUIBOR 6M':   '19,90%',
        'LUIBOR 9M':   '20,25%',
        'LUIBOR 12M':  '20,45%',
    },
    'luibor_d2': {
        'LUIBOR O/N':  '19,25%',
        'LUIBOR 1M':   '19,45%',
        'LUIBOR 3M':   '19,75%',
        'LUIBOR 6M':   '19,90%',
        'LUIBOR 9M':   '20,25%',
        'LUIBOR 12M':  '20,45%',
    },
    'luibor_variation': {
        'LUIBOR O/N':  '0,00%',
        'LUIBOR 1M':   '+0,05%',
        'LUIBOR 3M':   '0,00%',
        'LUIBOR 6M':   '+0,10%',
        'LUIBOR 9M':   '0,00%',
        'LUIBOR 12M':  '+0,05%',
    },
    'juros_diario_mn': '46,3 mM Kz',
    'fluxos_mn_rows': [
        {'label': 'Fluxos de Entradas (Cash in flow)',  'values': ['3.200', '2.800', '3.100', '2.900', '3.450']},
        {'label': 'Recebimentos de cupão de títulos',   'values': ['500',   '500',   '500',   '500',   '500']},
        {'label': 'Reembolsos de crédito (+)',           'values': ['1.200', '800',   '1.100', '900',   '1.200']},
        {'label': 'Reembolsos de OMA-O/N + Juros',      'values': ['1.500', '1.500', '1.500', '1.500', '1.750']},
        {'label': 'Fluxos de Saídas (Cash out flow)',   'values': ['2.800', '2.600', '2.950', '2.700', '3.050']},
        {'label': 'Desembolso de crédito (-)',           'values': ['1.800', '1.600', '1.900', '1.700', '2.000']},
        {'label': 'Aplicação em OMA',                   'values': ['1.000', '1.000', '1.050', '1.000', '1.050']},
        {'label': 'GAP de Liquidez',                    'values': ['+400',  '+200',  '+150',  '+200',  '+400']},
    ],
    'pl_summary': [
        {'label': 'Reembolso de Crédito', 'n_ops': 3, 'montante': '1.200 mM Kz'},
        {'label': 'Fornecedores',          'n_ops': 1, 'montante': '50 mM Kz'},
        {'label': 'Desembolso de Crédito', 'n_ops': 2, 'montante': '2.000 mM Kz'},
    ],
    'desembolsos_total': 2000,
    'reembolsos_pie': [
        {'label': 'BAI', 'valor': 600},
        {'label': 'BFA', 'valor': 400},
        {'label': 'BPC', 'valor': 200},
    ],
    'liquidez_me_rows': [
        {'label': 'SALDO D.O Estrangeiros', 'values': ['8,20', '8,50', '8,30', '8,60', '8,80']},
        {'label': 'DPs ME',                 'values': ['2,50', '2,50', '2,50', '2,50', '2,50']},
        {'label': 'COLATERAL CDI',           'values': ['1,30', '1,30', '1,30', '1,30', '1,30']},
        {'label': 'LIQUIDEZ BDA',            'values': ['12,00','12,30','12,10','12,40','12,60']},
    ],
    'transacoes_me_raw': [
        {'tipo': 'DP', 'moeda': 'USD', 'contraparte': 'Citibank', 'taxa': '4,50%', 'montante': '2,5 M USD', 'maturidade': '30/04/2026', 'juros': '0,03 M USD'},
    ],
    'operacoes_vivas_me': [
        {'contraparte': 'Citibank', 'montante': '2,5 M USD', 'taxa': '4,50%', 'residual': 29, 'vencimento': '30/04/2026', 'juro_diario': '0,031 M USD'},
        {'contraparte': 'ABSA',     'montante': '1,8 M USD', 'taxa': '4,25%', 'residual': 60, 'vencimento': '31/05/2026', 'juro_diario': '0,021 M USD'},
    ],
    'fluxos_me_rows': [
        {'label': 'Fluxos de entradas (Cash in flow)', 'values': ['0,50', '0,30', '0,40', '0,35', '0,45']},
        {'label': 'Outros recebimentos',               'values': ['0,10', '0,10', '0,10', '0,10', '0,10']},
        {'label': 'Reembolsos de DP + Juros',          'values': ['0,40', '0,20', '0,30', '0,25', '0,35']},
        {'label': 'Fluxos de Saídas (Cash out flow)',  'values': ['0,40', '0,25', '0,35', '0,28', '0,40']},
        {'label': 'Aplicação em DP ME',                'values': ['0,40', '0,25', '0,35', '0,28', '0,40']},
        {'label': 'GAP de Liquidez',                   'values': ['+0,10', '+0,05', '+0,05', '+0,07', '+0,05']},
    ],
    'juros_diario_me': '0,052 M USD',
    'cambial': {
        'vol_total_usd':  '12,5 M USD',
        'posicao_cambial': '5.230 mM Kz',
        'activos_usd':  45.2,
        'passivos_usd': 32.8,
    },
    'cambial_rows': [
        {'par': 'USD/AKZ', 'anterior2': '920.50', 'anterior': '921.00', 'atual': '922.30', 'variacao': '+0,14%'},
        {'par': 'EUR/AKZ', 'anterior2': '1010.20','anterior': '1012.00','atual': '1015.50','variacao': '+0,35%'},
        {'par': 'EUR/USD', 'anterior2': '1,098',  'anterior': '1,099',  'atual': '1,101',  'variacao': '+0,18%'},
    ],
    'transacoes_bda_rows': [
        {'cv': 'C', 'par': 'USD/AKZ', 'montante': '5,0 M USD', 'cambio': '922,30', 'pl': '+120 mM Kz'},
        {'cv': 'V', 'par': 'USD/AKZ', 'montante': '7,5 M USD', 'cambio': '922,50', 'pl': '+180 mM Kz'},
    ],
    'mercado_rows': [
        {'label': 'T+0', 'montante': '12,5 M USD', 'min': '920,00', 'max': '923,00'},
        {'label': 'T+1', 'montante': '8,2 M USD',  'min': '921,50', 'max': '922,80'},
        {'label': 'T+2', 'montante': '3,1 M USD',  'min': '921,00', 'max': '923,50'},
    ],
    'bodiva_segment_rows': [
        {'segmento': 'Obrigações De Tesouro',    'anterior': '45.230 mM Kz', 'atual': '45.890 mM Kz', 'variacao': '+1,46%'},
        {'segmento': 'Bilhetes Do Tesouro',       'anterior': '12.100 mM Kz', 'atual': '12.350 mM Kz', 'variacao': '+2,07%'},
        {'segmento': 'Obrigações Privadas',       'anterior': '3.200 mM Kz',  'atual': '3.210 mM Kz',  'variacao': '+0,31%'},
        {'segmento': 'Unidades De Participações', 'anterior': '1.500 mM Kz',  'atual': '1.520 mM Kz',  'variacao': '+1,33%'},
        {'segmento': 'Acções',                    'anterior': '850 mM Kz',    'atual': '870 mM Kz',    'variacao': '+2,35%'},
        {'segmento': 'Repos',                     'anterior': '200 mM Kz',    'atual': '195 mM Kz',    'variacao': '-2,50%'},
        {'segmento': 'Total',                     'anterior': '63.080 mM Kz', 'atual': '64.035 mM Kz', 'variacao': '+1,51%'},
    ],
    'bodiva_total_transacoes': '64.035 mM Kz',
    'bodiva_stocks': {
        'REFINA': {'volume': 15000, 'previous': 1850, 'current': 1900, 'change_pct': 2.70,  'cap_bolsista': '28.500 mM Kz'},
        'SEMBA':  {'volume': 8000,  'previous': 950,  'current': 960,  'change_pct': 1.05,  'cap_bolsista': '7.680 mM Kz'},
        'FINA':   {'volume': 5000,  'previous': 1200, 'current': 1185, 'change_pct': -1.25, 'cap_bolsista': '5.925 mM Kz'},
    },
    'bodiva_operacoes': [
        {'tipo': 'OT 2028', 'data': '01/04/2026', 'cv': 'C', 'preco': '98,50', 'quantidade': '500.000', 'montante': '492,5 mM Kz'},
        {'tipo': 'BT 364D', 'data': '01/04/2026', 'cv': 'V', 'preco': '91,20', 'quantidade': '200.000', 'montante': '182,4 mM Kz'},
    ],
    'bodiva_transacoes_valor': '674,9 mM Kz',
    'bodiva_juros_diario': '18,5 mM Kz',
    'carteira_titulos': [
        {'carteira': 'Custo Amortizado', 'cod': 'OT 2028', 'qty_d1': '500.000', 'qty_d': '500.000', 'nominal': '500 mM Kz', 'taxa': '9,50%',  'montante': '492,5 mM Kz', 'juros_anual': '46,8 mM Kz', 'juro_diario': '0,128 mM Kz'},
        {'carteira': 'Custo Amortizado', 'cod': 'OT 2030', 'qty_d1': '300.000', 'qty_d': '300.000', 'nominal': '300 mM Kz', 'taxa': '10,00%', 'montante': '298,0 mM Kz', 'juros_anual': '29,8 mM Kz', 'juro_diario': '0,082 mM Kz'},
        {'carteira': 'Justo Valor',      'cod': 'BT 364D', 'qty_d1': '200.000', 'qty_d': '200.000', 'nominal': '200 mM Kz', 'taxa': '18,50%', 'montante': '182,4 mM Kz', 'juros_anual': '33,7 mM Kz', 'juro_diario': '0,092 mM Kz'},
        {'carteira': 'Justo Valor',      'cod': 'OPr ABC', 'qty_d1': '100.000', 'qty_d': '100.000', 'nominal': '100 mM Kz', 'taxa': '11,00%', 'montante': '99,0 mM Kz',  'juros_anual': '10,9 mM Kz', 'juro_diario': '0,030 mM Kz'},
    ],
    'market_info': {
        'capital_markets': [
            {'indice': 'S&P 500',         'anterior': '5.254,35', 'atual': '5.218,19', 'variacao': '-0,69%'},
            {'indice': 'Dow Jones',        'anterior': '39.807',   'atual': '39.737',   'variacao': '-0,18%'},
            {'indice': 'NASDAQ',           'anterior': '16.379',   'atual': '16.156',   'variacao': '-1,36%'},
            {'indice': 'NIKKEI 225',       'anterior': '35.617',   'atual': '35.327',   'variacao': '-0,81%'},
            {'indice': 'IBOVESPA',         'anterior': '128.542',  'atual': '129.203',  'variacao': '+0,51%'},
            {'indice': 'EUROSTOXX 50',     'anterior': '4.892,23', 'atual': '4.876,45', 'variacao': '-0,32%'},
            {'indice': 'Bolsa de Londres', 'anterior': '7.932,10', 'atual': '7.958,30', 'variacao': '+0,33%'},
            {'indice': 'PSI 20',           'anterior': '6.712,45', 'atual': '6.698,20', 'variacao': '-0,21%'},
        ],
        'cm_commentary': 'Mercados globais encerram a semana em queda, pressionados por dados macroeconómicos nos EUA abaixo do esperado e incerteza sobre a trajectória das taxas de juro da Fed.',
        'crypto': [
            {'moeda': 'BITCOIN (BTC)',  'anterior': '71.230', 'atual': '69.845', 'variacao': '-1,94%'},
            {'moeda': 'ETHEREUM (ETH)', 'anterior': '3.512',  'atual': '3.480',  'variacao': '-0,91%'},
            {'moeda': 'XRP (XRP)',      'anterior': '0,612',  'atual': '0,608',  'variacao': '-0,65%'},
        ],
        'crypto_commentary': 'Criptomoedas registam correcção moderada, com Bitcoin a ceder abaixo dos 70.000 USD.',
        'commodities': [
            {'nome': 'PETRÓLEO (BRENT)',        'anterior': '87,42', 'atual': '85,90', 'variacao': '-1,74%'},
            {'nome': 'MILHO (USD/BU)',           'anterior': '4,32',  'atual': '4,28',  'variacao': '-0,93%'},
            {'nome': 'SOJA (USD/BU)',            'anterior': '11,85', 'atual': '11,92', 'variacao': '+0,59%'},
            {'nome': 'TRIGO (USD/LBS)',          'anterior': '5,91',  'atual': '5,88',  'variacao': '-0,51%'},
            {'nome': 'CAFÉ (USD/LBS)',           'anterior': '2,31',  'atual': '2,38',  'variacao': '+3,03%'},
            {'nome': 'AÇÚCAR (USD/LBS)',         'anterior': '0,198', 'atual': '0,195', 'variacao': '-1,52%'},
            {'nome': 'ÓLEO DE PALMA (USD/LBS)', 'anterior': '0,412', 'atual': '0,418', 'variacao': '+1,46%'},
            {'nome': 'ALGODÃO (USD/LBS)',        'anterior': '0,875', 'atual': '0,871', 'variacao': '-0,46%'},
            {'nome': 'BANANA (USD/LBS)',         'anterior': '0,320', 'atual': '0,325', 'variacao': '+1,56%'},
        ],
        'commodities_nota': 'O preço do Petróleo BRENT recuou 1,74% face ao dia anterior, pressionado pelas expectativas de aumento de produção da OPEP+ e dados de inventários dos EUA acima do esperado.',
        'minerais': [
            {'nome': 'OURO',     'anterior': '2.315,40', 'atual': '2.328,70', 'variacao': '+0,57%'},
            {'nome': 'FERRO',    'anterior': '110,50',   'atual': '109,80',   'variacao': '-0,63%'},
            {'nome': 'COBRE',    'anterior': '9.812,00', 'atual': '9.856,00', 'variacao': '+0,45%'},
            {'nome': 'MANGANÊS', 'anterior': '4,25',     'atual': '4,22',     'variacao': '-0,71%'},
        ],
        'minerais_commentary': 'Ouro mantém tendência de alta; metais industriais com ligeira correcção.',
    },
}


def build_sample_report(output_path='output/test_sample_01042026.pptx'):
    gen = BDAReportGenerator(data)
    return gen.build(output_path)


if __name__ == "__main__":
    path = build_sample_report()
    print(f"Generated: {path}")
