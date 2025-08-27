import pandas as pd
import numpy as np
import urllib.request, urllib.parse, urllib.error
from bs4 import BeautifulSoup
import ssl
ctx = ssl.create_default_context()
ctx.check_hostname = False
ctx.verify_mode = ssl.CERT_NONE

headers = {
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) '
                  'AppleWebKit/537.36 (KHTML, like Gecko) '
                  'Chrome/91.0.4472.124 Safari/537.36'
}

IS_statements = ["revenue", "ebitda", "amortization", "opex", "employment_expenses", "operating_income", "interest_expenses", "net_income"]
important_stats = ['revenue', 'ebitda', 'operating_income', 'net_income', 'capex', 'fcf', 'div_yield', 'assets', 'debt', 'cash', 'eps', 'ebitda_margin', 'net_margin', 'roe', 'roa', 'p_e', 'p_bv', 'ev_ebitda', 'debt_ebitda', 'capex_revenue']
years = ['2020', '2021', '2022', '2023', '2024']

def get_smartlab_statements(ticker, statements=important_stats, target_years=years, translation=True, types='MSFO', horizontal_an=False):
    '''
    Takes ticker of the company and turns it into URL from smart-lab.ru section of the given company and returns DataFrame of the desired statement
    ticker - string, company's official ticker
    statements - array of the strings of financial indicators
    years - array of the desired years
    translations_IS - array of the convinient translation of indexes
    types - a string (only MSFO ans RSBU are posssible)
    '''
    
    if types not in ('MSFO', 'RSBU'):
        return print('ERROR: entered type should be either MSFO or RSBU')
            
    try:   
        url = f'https://smart-lab.ru/q/{ticker}/f/y/{types}/'
        req = urllib.request.Request(url, headers=headers)
        with urllib.request.urlopen(req, context=ctx) as response:
            html = response.read()
        soup = BeautifulSoup(html, 'html.parser')
    except:
        return print('Ticker is invalid')

    # Определяем реальные годы по заголовку таблицы
    header_row = soup.find('tr', class_='header_row')
    header_cells = header_row.find_all('td')
    site_years = [cell.get_text(strip=True).replace('?', '') for cell in header_cells]
    
    fm = pd.DataFrame(index=statements, columns=target_years)
    translations = []
    
    # Обрабатываем каждый показатель
    for st in statements:
        row = soup.find('tr', {'field': st})
        if row is None:
            print(f'There is no {st.upper()} for {ticker} in {types}')
            fm = fm.drop(st)
            continue
        tds = row.find_all('td')
        ths = row.find('th').get_text(strip=True)
        translations.append(ths)
        values = []
    
        for td in tds:
            text = td.get_text(strip=True).replace('%', '').replace(' ', '')
            try:
                val = float(text)
            except ValueError:
                val = np.nan
            values.append(val)
    
        # Сопоставим значения с годами (только нужные годы)
        for year in target_years:
            if year in site_years:
                idx = site_years.index(year)
                if idx < len(values):
                    fm.loc[st, year] = values[idx]

    # Перевод 
    if translation:
        fm.index = translations  

    # Горизонтальный анализ
    if horizontal_an:
        hor_an_dt = fm.copy()
        hor_an_dt = hor_an_dt.drop(hor_an_dt.columns[0], axis=1)
        for row in hor_an_dt.index:
            for i in range(1, len(fm.columns)):
                year_prev = fm.columns[i - 1]
                year_curr = fm.columns[i]
                prev_val = float(fm.loc[row, year_prev])
                curr_val = float(fm.loc[row, year_curr])
        
                if pd.notna(prev_val) and prev_val != 0:
                    growth = (curr_val - prev_val) / prev_val * 100
                    hor_an_dt.loc[row, year_curr] = f'{round(growth, 2)}%'
                else:
                    hor_an_dt.loc[row, year_curr] = np.nan
        hor_an_dt.columns = [f"{col} изм (%)" for col in hor_an_dt.columns]
        fm = pd.concat([fm, hor_an_dt], axis=1)
           
    return fm
