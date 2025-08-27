import streamlit as st
import pandas as pd
import numpy as np
import re
import fin_an as fa
from io import BytesIO

av = pd.DataFrame(columns=['Сравнение', 'Средние значения'])
av.loc['Текущая ликвидность'] = ['>', 1.78]
av.loc['Быстрая ликвиднсоть'] = ['>', 1.26]
av.loc['Срочная ликвидность'] = ['>', 0.17]
av.loc['ROS, %'] = ['>', 5.7]
av.loc['EBIT Margin, %'] = ['>', 4.9]
av.loc['ROA, %'] = ['>', 6.8]
av.loc['ROE, %'] = ['>', 29.7]
av.loc['Пер. об. Активов'] = ['<', 183]
av.loc['Пер. об. Дебиторки'] = ['<', 62]
av.loc['Пер. об. Запасов'] = ['<', 16]
av.loc['Пер. об. Кредиторки'] = ['<', 90]
av.loc['Коэф. автономии'] = ['>', 0.37]
av.loc['Debt/Equity'] = ['<', 2]
av.loc['Gearing'] = ['<', 0.7]
av.loc['ICR'] = ['>', 5.96]
av.loc['Net WC (КОСОС)'] = ['>', 0.26]
av.loc['FCF Margin, %'] = ['>', 15]
av.loc['FCF / Net Income'] = ['>', 1] # Лучше поставить равно
av.loc['Cash Flow to Debt'] = ['>', 1.5]
av.loc['EPS,руб'] = ['>', 10]
av.loc['P/E'] = ['<', 7]
av.loc['P/S'] = ['<', 2]
av.loc['P/BV'] = ['<', 1] # Лучше поставить равно
av.loc['EV/EBITDA'] = ['<', 12]
av.loc['Долг/EBITDA'] = ['<', 2]

def make_unique(rows):
    seen = {}
    new_rows = []
    for row in rows:
        if row in seen:
            seen[row] += 1
            # добавляем N пробелов в зависимости от числа повторений
            new_rows.append(row + " " * seen[row])
        else:
            seen[row] = 0
            new_rows.append(row)
    return new_rows

def get_income_statement(file, dropna=False, analysis=False, excel_file=False):
    '''
    takes garbage БФО excel file and turns it into normal one
    '''
    df = pd.read_excel(file, sheet_name='Отчет о финансовых результатах', header=4, index_col=4, engine='openpyxl')
    pattern = r'^Код строки$|^За\s\d{4}\sг\.'
    df = df.filter(regex=pattern)
    df = df[df['Код строки'].notna()]
    df = df.drop('Код строки', axis=1)
    df.index = df.index.str.replace(r'\d+', '', regex=True).str.strip()
    df = df.loc['Выручка':'Чистая прибыль (убыток)']
    df = df.replace(' ', '', regex=True).replace(r'\(', '-', regex=True).replace(r'\)', '', regex=True)
    df = df.apply(pd.to_numeric, errors='coerce')

    if dropna:
        df = df.dropna()

    # Вертикальный и горизонтальный анализ
    if analysis:
        # Первый цикл берет за k кол-во столбцов - 1, чтобы вернуть года из их названия
        # и автоматизировать переход от столбца к столбцу
        for k in range(df.shape[1]):
            hor_an = []
            ver_an = []
            year1 = df.columns.str.extract(r'(\d{4})')[0][k]
            year0 = df.columns.str.extract(r'(\d{4})')[0][k+1]

            # Второй цикл берет за i кол-во строк, чтобы переходить от строки к строке
            # Проверка нужна, чтобы не брать несуществующие года для горизонтального анализа
            if k == 0:
                for i in range(df.shape[0]):
                    dif = np.round((df.iloc[i,k] / df.iloc[i,k+1] - 1) * 100, 2)
                    hor_an.append(dif)
                df[f'{year1} / {year0}'] = hor_an
                
            for j in range(df.shape[0]):
                frac = abs(np.round((df.iloc[j,k] / df.iloc[0,k]) * 100, 2))
                ver_an.append(frac)   
            
            df[f'Доля в выручке {year1}'] = ver_an

        # Стилизация
        # Делаем уникальные строки, чтобы Styler не ругался и находим столбцы анализа, чтобы форматировать только их. 
        # Также находим оригинальные столбцы и столбцы для вертикального анализа, чтобы их нормально отформатировать
        df.index = make_unique(df.index)
        hor_analysis_cols = [col for col in df.columns if "/" in col]
        ver_analysis_cols = [col for col in df.columns if "Доля" in col]
        orig_cols = [col for col in df.columns if "За" in col]

        
        def color_vals(val):
            if isinstance(val, (int, float, np.number)):
                if val > 0:
                    return "color: green"
                elif val < 0:
                    return "color: red"
            return ""
            
        # Чтобы один формат не перебивал другой
        formats = {col: "{:.2f}%" for col in hor_analysis_cols}
        formats.update({col: "{:.2f}%" for col in ver_analysis_cols})
        formats.update({col: "{:,.0f}" for col in orig_cols})
        
        return (
            df.style
              .applymap(color_vals, subset=hor_analysis_cols)
              .format(formats)
        ), None

    # Скачивать ли эксель файл
    if excel_file:
        output = BytesIO()
        df.to_excel(output)
        output.seek(0)
        return df, output
            
    return df, None

    

def get_balance(file, dropna=False, analysis=False, excel_file=False):
    '''
    takes garbage БФО excel file and turns it into normal one,
    gives you only necessary items, the extra ones are excluded
    '''
    df = pd.read_excel(file, sheet_name='Бухгалтерский баланс', header=4, index_col=3, engine='openpyxl')
    pattern = r'^Код строки$|^На \d{1,2} [а-яё]+ \d{4} г\.$'
    df = df.filter(regex=pattern, axis=1)
    df = df[df['Код строки'].notna()]
    df = df.drop('Код строки', axis=1)
    df = df.drop(df.index[0])
    df = df.replace(' ', '', regex=True).replace(r'\(', '-', regex=True).replace(r'\)', '', regex=True)
    df = df.apply(pd.to_numeric, errors='coerce')

    if dropna:
        df = df.dropna()

    # Вертикальный и горизонтальный анализ
    if analysis:
        # Первый цикл берет за k кол-во столбцов - 1, чтобы вернуть года из их названия
        # и автоматизировать переход от столбца к столбцу
        for k in range(df.shape[1]):
            hor_an = []
            ver_an = []
            year1 = df.columns.str.extract(r'(\d{4})')[0][k]
            year0 = df.columns.str.extract(r'(\d{4})')[0][k+1]

            # Второй цикл берет за i кол-во строк, чтобы переходить от строки к строке
            # Проверка нужна, чтобы не брать несуществующие года для горизонтального анализа
            if k < 2:
                for i in range(df.shape[0]):
                    dif = np.round((df.iloc[i,k] / df.iloc[i,k+1] - 1) * 100, 2)
                    hor_an.append(dif)
                df[f'{year1} / {year0}'] = hor_an
                
            for j in range(df.shape[0]):
                frac = np.round((df.iloc[j,k] / df.loc['БАЛАНС'].iloc[0,k]) * 100, 2)
                ver_an.append(frac) 
                
            df[f'Доля в балансе {year1}'] = ver_an

        # Делаем нормальный порядок столбцов
        cols = df.columns.tolist()
        cols[4], cols[5] = cols[5], cols[4]
        df = df[cols]

        # Начало стилизации
        # Делаем уникальные строки, чтобы Styler не ругался и находим столбцы анализа, чтобы форматировать только их. 
        # Также находим оригинальные столбцы и столбцы для вертикального анализа, чтобы их нормально отформатировать
        df.index = make_unique(df.index)
        hor_analysis_cols = [col for col in df.columns if "/" in col]
        ver_analysis_cols = [col for col in df.columns if "Доля" in col]
        orig_cols = [col for col in df.columns if "На" in col]

        def color_vals(val):
            if isinstance(val, (int, float, np.number)):
                if val > 0:
                    return "color: green"
                elif val < 0:
                    return "color: red"
            return ""
            
        # Чтобы один формат не перебивал другой
        formats = {col: "{:.2f}%" for col in hor_analysis_cols}
        formats.update({col: "{:.2f}%" for col in ver_analysis_cols})
        formats.update({col: "{:,.0f}" for col in orig_cols})
        
        return (
            df.style
              .applymap(color_vals, subset=hor_analysis_cols)
              .format(formats)
        ), None

    # Скачивать ли эксель файл
    if excel_file:
        output = BytesIO()
        df.to_excel(output)
        output.seek(0)
        return df, output

    return df, None



def get_cash_flow_statement(file, only_OCF_FCF=False, dropna=False, analysis=False, excel_file=False):
    '''
    takes garbage БФО excel file and turns it into normal one
    '''
    try:
        df = pd.read_excel(file, sheet_name='Отчет о движении денежных средс', header=4, index_col=0, engine='openpyxl')
    except:
        return print('Такой отчетности нет в файле')
    pattern = r'За\s\d{4}\sг\.'
    df = df.filter(regex=pattern, axis=1)
    df.index = df.index.str.replace('в том числе:\n ', '', regex=True).str.replace(r'\s+', ' ', regex=True).str.replace('4127. ', '', regex=False).str.strip()
    df = df.drop(df.index[0])
    df = df[~df.apply(lambda row: all(val in ['-', '(-)'] for val in row), axis=1)]
    df = df.replace(' ', '', regex=True).replace(r'\(', '-', regex=True).replace(r'\)', '', regex=True)
    df = df.apply(pd.to_numeric, errors='coerce')

    # Возвращает только OCF и FCF
    if only_OCF_FCF:
        OCF = df.loc['Сальдо денежных потоков от текущих операций']
        try:
            FCF = OCF + df.loc['в связи с приобретением, созданием, модернизацией, реконструкцией и подготовкой к использованию внеоборотных активов']
        except:
            FCF = OCF
        return OCF, FCF
        
    if dropna:
        df = df.dropna()   

    # Вертикальный и горизонтальный анализ
    if analysis:
        # Первый цикл берет за k кол-во столбцов - 1, чтобы вернуть года из их названия
        # и автоматизировать переход от столбца к столбцу
        
        hor_an = []
        year1 = df.columns.str.extract(r'(\d{4})')[0][0]
        year0 = df.columns.str.extract(r'(\d{4})')[0][1]

        # Цикл берет за i кол-во строк, чтобы переходить от строки к строке
        for i in range(df.shape[0]):
            dif = np.round((df.iloc[i,0] / df.iloc[i,1] - 1) * 100, 2)
            hor_an.append(dif)
                
        df[f'{year1} / {year0}'] = hor_an

        # Стилизация
        # Делаем уникальные строки, чтобы Styler не ругался и находим столбцы анализа, чтобы форматировать только их. 
        # Также находим оригинальные столбцы и столбцы для горизонтального анализа, чтобы их нормально отформатировать
        df.index = make_unique(df.index)
        hor_analysis_cols = [col for col in df.columns if "/" in col]
        orig_cols = [col for col in df.columns if col not in hor_analysis_cols]
 
        def color_vals(val):
            if isinstance(val, (int, float, np.number)):
                if val > 0:
                    return "color: green"
                elif val < 0:
                    return "color: red"
            return ""
            
        # Чтобы один формат не перебивал другой
        formats = {col: "{:.2f}%" for col in hor_analysis_cols}
        formats.update({col: "{:,.0f}" for col in orig_cols})
        
        return (
            df.style
              .applymap(color_vals, subset=hor_analysis_cols)
              .format(formats)
        ), None

    # Скачивать ли эксель файл
    if excel_file:
        output = BytesIO()
        df.to_excel(output)
        output.seek(0)
        return df, output

    return df, None



def get_smartlab_ratios(ticker, statements_RSBU=['eps', 'p_e', 'p_s', 'p_bv'], statements_MSFO=['ev_ebitda', 'debt_ebitda'], years=['2024', '2023']):
    '''
    searches for required ratios of the given company on smartlab (both RSBU and MSFO pages)
    and returns a DataFrame with them
    '''
    RSBU = fa.get_smartlab_statements(ticker, statements=statements_RSBU, target_years=years, types='RSBU', horizontal_an=False)
    MSFO = fa.get_smartlab_statements(ticker, statements=statements_MSFO, target_years=years, types='MSFO', horizontal_an=False)
    df = pd.concat([RSBU, MSFO], axis=0)
    return df


def get_ratios(IS, balance, OCF=None, FCF=None, smartlab_df=None, extra_ratios=True, styled=False):
    balance = balance.drop(balance.columns[-1], axis=1)
    balance.columns = IS.columns
    ratios = pd.DataFrame(columns=balance.columns)
    ratios.loc['Текущая ликвидность'] = balance.loc['Итого по разделу II'] / balance.loc['Итого по разделу V']
    ratios.loc['Быстрая ликвиднсоть'] = (balance.loc['Итого по разделу II'] - balance.loc['Запасы']) / balance.loc['Итого по разделу V']
    ratios.loc['Срочная ликвидность'] = (balance.loc['Денежные средства и денежные эквиваленты'] + balance.loc['Финансовые вложения (за исключением денежных эквивалентов)']) / balance.loc['Итого по разделу V']
    ratios.loc['ROS, %'] = 100 * IS.loc['Чистая прибыль (убыток)'] / IS.loc['Выручка']
    ratios.loc['EBIT Margin, %'] = 100 * IS.loc['Прибыль (убыток) от продаж'] / IS.loc['Выручка']
    ratios.loc['ROA, %'] = 100 * IS.loc['Чистая прибыль (убыток)'] / balance.loc['БАЛАНС'].iloc[0]
    ratios.loc['ROE, %'] = 100 * IS.loc['Чистая прибыль (убыток)'] / balance.loc['Итого по разделу III']
    ratios.loc['Пер. об. Активов'] = 365 * balance.loc['БАЛАНС'].iloc[0] / IS.loc['Выручка']
    ratios.loc['Пер. об. Дебиторки'] = 365 * balance.loc['Дебиторская задолженность'] / IS.loc['Выручка']
    ratios.loc['Пер. об. Запасов'] = -365 * balance.loc['Запасы'] / IS.loc['Себестоимость продаж']
    ratios.loc['Пер. об. Кредиторки'] = -365 * balance.loc['Кредиторская задолженность'] / IS.loc['Себестоимость продаж']
    ratios.loc['Коэф. автономии'] = balance.loc['Итого по разделу III'] / balance.loc['БАЛАНС'].iloc[0]
    ratios.loc['Debt/Equity'] = balance.loc['Заемные средства'].sum(axis=0) / balance.loc['Итого по разделу III']
    ratios.loc['Gearing'] = balance.loc['Заемные средства'].sum(axis=0) / (balance.loc['Итого по разделу III'] + balance.loc['Заемные средства'].sum(axis=0))
    ratios.loc['ICR'] = IS.loc['Прибыль (убыток) от продаж'] / -IS.loc['Проценты к уплате']
    ratios.loc['Net WC (КОСОС)'] = (balance.loc['Итого по разделу II'] - balance.loc['Итого по разделу V']) / balance.loc['Итого по разделу II']
    # Можно еще добавить DSCR (инфа есть в гпт)

    if extra_ratios:
        ratios.loc['Cash Flow to Debt'] = OCF / balance.loc['Заемные средства'].sum(axis=0)
        ratios.loc['FCF Margin, %'] = 100 * FCF / IS.loc['Выручка']
        ratios.loc['FCF / Net Income'] = FCF / IS.loc['Чистая прибыль (убыток)']
        try:
            smartlab_df.columns = IS.columns
            ratios = pd.concat([ratios, smartlab_df], axis=0) 
            ratios = ratios.apply(pd.to_numeric, errors='coerce')
        except: 
            ratios = ratios

    ratios = ratios.round(2)
    
    common_index = ratios.index.intersection(av.index)
    ratios = pd.concat([ratios.loc[common_index], av.loc[common_index]], axis=1)

    def highlight_compare(row):
        result = []
    
        for col in range(ratios.shape[1] - 2):
            value = row[col]
            sign = row['Сравнение']
            avg = row['Средние значения']
    
            if sign == '>' and value > avg:
                result.append('color: green')
            elif sign == '>' and value < avg:
                result.append('color: red')
            elif sign == '<' and value < avg:
                result.append('color: green')
            elif sign == '<' and value > avg:
                result.append('color: red')
            else:
                result.append('')
        
        result.extend(['', ''])  
        return result

        # Применением при необходимости
    if styled:
        style = ratios.style\
            .apply(highlight_compare, axis=1)\
            .format({col: '{:,.2f}' for col in ratios.select_dtypes('number').columns})
        return style
        
    return ratios


# Функции Streamlit
# Ключи этого словаря используются в выпадающем меню, а значения - в качестве функций
names = {"ОФР": get_income_statement, "Баланс": get_balance, "ОДДС": get_cash_flow_statement}
def show_statements(report_type):
    drop_na = st.checkbox("Убрать строчки с NaN")
    do_analysis = st.checkbox("Выполнить горизонтальный и вертикальный анализ")
    dowload_excel = st.checkbox("Скачать файл Excel")
    df, excel_file = names[report_type](selected_file, dropna=drop_na, analysis=do_analysis, excel_file=dowload_excel)
    if dowload_excel and excel_file:
        st.download_button(
        label="Скачать Excel",
        data=excel_file,
        file_name="ОФР_чистый.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    st.subheader(f'{report_type}')
    st.dataframe(df, use_container_width=True)


# Интерфейс Streamlit
st.set_page_config(
    page_title="Финансовый Анализ",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.title("Обработка отчетности БФО")
uploaded_files = st.file_uploader("Загрузите Excel-файлы", type=["xlsx"], accept_multiple_files=True)

if uploaded_files:
    file_names = [file.name for file in uploaded_files]
    selected_file_name = st.selectbox("Выберите файл для анализа", file_names)

    selected_file = next(file for file in uploaded_files if file.name == selected_file_name)

    mode = st.radio("Выберите режим анализа:", ["Анализ отчетности", "Финансовые коэффициенты"])

    if mode == "Финансовые коэффициенты":
        company_type = st.radio("Выберите вид компании:", ["Публичная", "Непубличная"])
        st.subheader("📌 Финансовые коэффициенты")
        style = st.checkbox('Условное форматирование')

        # Преобразуем отчетности из файла
        try:
            ofr_df, a = get_income_statement(selected_file)
            balance_df, b = get_balance(selected_file)

            # Передаём уже готовые таблицы в функцию расчёта
            if company_type == 'Публичная':
                ticker = st.text_input("Введите тикер компании, чтобы получить доп коэффициенты со smartlab")
                ocf, fcf = get_cash_flow_statement(selected_file, only_OCF_FCF=True)
                # Проверка на нормальный тикер
                try:
                    # Для того, чтобы при пустом поле ничего не отображалось
                    smartlab_df = get_smartlab_ratios(ticker) if ticker else None
                    ratios_df = get_ratios(ofr_df, balance_df, OCF=ocf, FCF=fcf, smartlab_df=smartlab_df, styled=style)  
                except:
                    st.warning("Введен несуществующий тикер")
                    ratios_df = get_ratios(ofr_df, balance_df, OCF=ocf, FCF=fcf, smartlab_df=None, styled=style)
            else:
                ratios_df = get_ratios(ofr_df, balance_df, extra_ratios=False, styled=style)

            st.dataframe(ratios_df, use_container_width=True)


        except Exception as e:
            st.error(f"Ошибка при расчете коэффициентов (возможно, загрузили неправильный файл или выбрали неправильный вид компании): {e}")

    else:
        report_type = st.selectbox("Выберите тип отчетности:", names.keys())
        show_statements(report_type=report_type)

