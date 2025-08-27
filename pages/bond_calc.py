import pandas as pd
import numpy as np
import streamlit as st

def discount(t, r):
    return pd.Series([(1+r)**-i for i in t], index=t)

def pv(flows, r):
    discounts = discount(flows.index, r)
    return (discounts * flows).sum()

def inst_to_ann(r):
    return np.expm1(r)

def ann_to_inst(r):
    return np.log1p(r)

def bond_cash_flows(maturity=10, principal=1000, coupon_rate=0.03, coupons_per_year=12):
    n_coupons = maturity*coupons_per_year
    coupon_amt = principal*coupon_rate/coupons_per_year
    coupon_times = np.arange(1, n_coupons+1)
    cash_flows = pd.Series(data=coupon_amt, index=coupon_times)
    cash_flows.iloc[-1] += principal
    return cash_flows

def bond_price(maturity=10, principal=1000, coupon_rate=0.03, coupons_per_year=12, discount_rate=0.03):
    if isinstance(discount_rate, pd.DataFrame):
        pricing_dates = discount_rate.index
        prices = pd.DataFrame(index = pricing_dates, columns = discount_rate.columns)
        for t in pricing_dates:
            prices.loc[t] = bond_price(maturity-t/coupons_per_year, principal, coupon_rate, coupons_per_year, discount_rate.loc[t])
        return prices
    else:
        if maturity <= 0: return principal+principal*coupon_rate/coupons_per_year
        cash_flows = bond_cash_flows(maturity, principal, coupon_rate, coupons_per_year)
        return pv(cash_flows, discount_rate/coupons_per_year)

def bond_current_yield(principal=1000, current_price=1000, coupon_rate=0.2):
    annual_coupon_payment = coupon_rate * principal
    return annual_coupon_payment/current_price

def bond_ytm(maturity=10, principal=1000, current_price=950, coupon_rate=0.2):
    """Пока очень упрощенная, без разных купонов, цен и тд"""
    annual_coupon_payment = coupon_rate * principal
    return (annual_coupon_payment+(principal-current_price)/maturity) / ((principal+current_price)/2)
        
def bond_total_return(monthly_prices, principal, coupon_rate, coupons_per_year):
    coupons = pd.DataFrame(data = 0, index=monthly_prices.index, columns=monthly_prices.columns)
    t_max = monthly_prices.index.max()
    pay_date = np.linspace(12/coupons_per_year, t_max, int(coupons_per_year*t_max/12), dtype=int)
    coupons.iloc[pay_date] = principal*coupon_rate/coupons_per_year
    total_returns = (monthly_prices + coupons)/monthly_prices.shift()-1
    return total_returns.dropna()

def macaulay_duration(flows, discount_rate):
    discounts = discount(flows.index, discount_rate)
    discounted_flows = flows * discounts
    weights = discounted_flows / discounted_flows.sum()
    return np.average(flows.index, weights=weights) / 12

def match_durations(cf_t, cf_s, cf_l, discount_rate):
    d_t = macaulay_duration(cf_t, discount_rate)
    d_s = macaulay_duration(cf_s, discount_rate)
    d_l = macaulay_duration(cf_l, discount_rate)
    return (d_l - d_t)/(d_l - d_s)

def bond_summary(name='Облигация', principal=1000, maturity=10, current_price=1000, coupon_rate=0.03, coupons_per_year=12, discount_rate=0.03):
    flows = bond_cash_flows(maturity, principal, coupon_rate, coupons_per_year)
    market_price = bond_price(maturity, principal, coupon_rate, coupons_per_year, discount_rate)
    current_yield = bond_current_yield(principal, current_price, coupon_rate)
    ytm = bond_ytm(maturity, principal, current_price, coupon_rate)
    mac_dur = macaulay_duration(flows, ytm/coupons_per_year)
    return pd.DataFrame({
        "Market Price": market_price,
        "Current Yield": current_yield,
        "YTM": ytm,
        "Macaulay Duration": mac_dur
    }, index = [name])

# Функции для streamlit

def add_bond(summary_df):
    """Добавляет рассчитанную облигацию в общий портфель (session_state)."""
    if "bond_portfolio" not in st.session_state:
        st.session_state["bond_portfolio"] = pd.DataFrame()

    st.session_state["bond_portfolio"] = pd.concat(
        [st.session_state["bond_portfolio"], summary_df]
    )

def clear_portfolio():
    """Полностью очищает портфель облигаций."""
    st.session_state["bond_portfolio"] = pd.DataFrame()

def remove_bond(name):
    """Удаляет облигацию из портфеля по имени."""
    if "bond_portfolio" in st.session_state and not st.session_state["bond_portfolio"].empty:
        st.session_state["bond_portfolio"] = st.session_state["bond_portfolio"].drop(index=name, errors="ignore")


# Тест облигаций
st.title("📉 Анализ облигаций")

st.sidebar.header("Параметры облигации")

# Ввод параметров пользователем
name = st.sidebar.text_input("Название облигации", "Облигация АО")
principal = st.sidebar.number_input("Номинал (₽)", min_value=100, max_value=10_000_000, value=1000, step=100)
maturity = st.sidebar.number_input("Срок до погашения (лет)", min_value=0.0, max_value=50.0, value=2.0, step=0.01)
current_price = st.sidebar.number_input("Текущая цена (₽)", min_value=100, max_value=10_000_000, value=1000, step=1)
coupon_rate = st.sidebar.number_input("Купонная ставка (в долях, 0.03 = 3%)", min_value=0.0, max_value=1.0, value=0.2, step=0.001, format="%.3f")
coupons_per_year = st.sidebar.selectbox("Число выплат в год", [1, 2, 4, 12], index=3)
discount_rate = st.sidebar.number_input("Дисконтная ставка (в долях)", min_value=0.0, max_value=1.0, value=0.18, step=0.001, format="%.3f")

calc = st.sidebar.button('Рассчитать')

# Кнопка расчёта
if calc:
    try:
        summary = bond_summary(
            name=name,
            principal=principal,
            maturity=maturity,
            current_price=current_price,
            coupon_rate=coupon_rate,
            coupons_per_year=coupons_per_year,
            discount_rate=discount_rate
        )
        st.session_state["last_summary"] = summary
        if "last_summary" in st.session_state:
            add_bond(st.session_state["last_summary"])

    except Exception as e:
        st.error(f"Ошибка при расчете: {e}")


# Управление портфелем
if "bond_portfolio" in st.session_state and not st.session_state["bond_portfolio"].empty:
    st.subheader("📊 Результаты анализа")
    st.dataframe(st.session_state["bond_portfolio"].style.format("{:,.4f}"))

    c1, c2 = st.columns(2)

    # Очистить портфель
    if c1.button("🗑 Очистить все", key="clear"):
        clear_portfolio()
        st.rerun()  # сразу перерисовываем страницу

    # --- Удаление отдельных облигаций через multiselect ---
    bonds_to_remove = c2.multiselect(
        "Удалить облигацию:",
        st.session_state["bond_portfolio"].index.tolist(),
        key="remove"
    )

    if bonds_to_remove:
        st.session_state["bond_portfolio"] = st.session_state["bond_portfolio"].drop(index=bonds_to_remove)
        st.rerun()

