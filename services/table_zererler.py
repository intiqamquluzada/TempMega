import pandas as pd
from datetime import datetime
import numpy as np

dz = pd.read_excel('/Users/quluzade/Desktop/mega-prod/data_files/ZJ 5 illik.xlsx', engine='openpyxl')
input_date_str = input("Tarixi YYYY-MM-DD formatında daxil edin: ")
input_date = datetime.strptime(input_date_str, "%Y-%m-%d")


def generate_quarterly_range(input_date):
    end_date = last_day_of_quarter(input_date)
    start_date = end_date - pd.DateOffset(years=3) + pd.DateOffset(days=1)

    quarters = pd.date_range(start=start_date, end=end_date, freq='Q')
    quarter_end_dates = quarters.map(last_day_of_quarter)

    return quarter_end_dates


def last_day_of_quarter(Tarix):
    quarter = (Tarix.month - 1) // 3 + 1
    if quarter == 1:
        return pd.Timestamp(year=Tarix.year, month=3, day=31)
    elif quarter == 2:
        return pd.Timestamp(year=Tarix.year, month=6, day=30)
    elif quarter == 3:
        return pd.Timestamp(year=Tarix.year, month=9, day=30)
    elif quarter == 4:
        return pd.Timestamp(year=Tarix.year, month=12, day=31)


def form3(data, input_date):
    quarterly_dates = generate_quarterly_range(input_date)
    filtered_df = data[data['SigortaHadisesininBasverdiyiTarix'].apply(last_day_of_quarter).isin(quarterly_dates) &
                       data['VerilmisOdenisTarixi'].isna() & data['VerilmisSigortaOdenisi'].isna()]
    sorted_df = filtered_df.sort_values('SigortaHadisesininBasverdiyiTarix').reset_index(drop=True)

    unique_quarters = sorted(sorted_df['SigortaHadisesininBasverdiyiTarix'].apply(last_day_of_quarter).unique())
    new_df = pd.DataFrame(index=[''] + unique_quarters)
    new_df['Sətrin kodu'] = [f'{i:02d}' for i in range(1, len(new_df) + 1)]
    new_df['Bildirilmiş, lakin tənzimlənməmiş zərərlərin hesabat tarixinə məbləği'] = \
    sorted_df.groupby(sorted_df['SigortaHadisesininBasverdiyiTarix'].apply(last_day_of_quarter))[
        'SigortaOdenisiUzreCemiBorc'].sum().reindex(new_df.index, fill_value=0)
    new_df.iloc[0, 1] += data[(data['SigortaHadisesininBasverdiyiTarix'] <= quarterly_dates[0]) &
                              data['VerilmisOdenisTarixi'].isna() & data['VerilmisSigortaOdenisi'].isna()][
        'SigortaOdenisiUzreCemiBorc'].sum()

    new_df['Sığorta müqavilələrinə vaxtından əvvəl xitam verilməsi məbləği'] = 0
    new_df['Zərərlərin tənzimlənməsi xərcləri (III+ IV)x0,03'] = new_df.iloc[:, 1] * 0.03
    new_df['Bildirilmiş zərərlər ehtiyatı (III+IV+V)'] = new_df.iloc[:, 1:4].sum(axis=1)

    new_df.loc['Yekun BTZE'] = ['X'] + list(new_df.iloc[:, 1:].sum())
    new_df.index = [''] + [f'rüb {i}' for i in range(len(unique_quarters), 0, -1)] + ['Yekun BTZE']
    return new_df.reset_index().rename(columns={'index': 'Sığorta hadisələrinin baş verdiyi rüblər'})


import pandas as pd


def shift_rows_left(df):
    shifted_df = pd.DataFrame(index=df.index, columns=df.columns)
    index_positions = range(len(df))

    for idx, pos in zip(df.index, index_positions):
        shift_amount = pos
        shifted_row = [0] * len(df.columns)
        row_data = df.loc[idx].tolist()
        if shift_amount < len(row_data):
            shifted_row[:len(row_data) - shift_amount] = row_data[shift_amount:]
        shifted_df.loc[idx] = shifted_row

    shifted_df.columns = df.columns
    return shifted_df


def last_day_of_quarter(Tarix):
    quarter = (Tarix.month - 1) // 3 + 1
    if quarter == 1:
        return pd.Timestamp(year=Tarix.year, month=3, day=31)
    elif quarter == 2:
        return pd.Timestamp(year=Tarix.year, month=6, day=30)
    elif quarter == 3:
        return pd.Timestamp(year=Tarix.year, month=9, day=30)
    elif quarter == 4:
        return pd.Timestamp(year=Tarix.year, month=12, day=31)


def generate_quarterly_range(input_date):
    end_date = last_day_of_quarter(input_date)
    start_date = end_date - pd.DateOffset(years=3)

    quarters = pd.date_range(start=start_date, end=end_date, freq='Q')
    quarter_end_dates = quarters.map(last_day_of_quarter)

    return quarter_end_dates


def form8(df, input_date):
    quarterly_dates = generate_quarterly_range(input_date)

    df['zerer_kvartal'] = df['SigortaHadisesininBasverdiyiTarix'].apply(last_day_of_quarter)
    df['odenis_kvartal'] = df['VerilmisOdenisTarixi'].apply(last_day_of_quarter)

    filtered_df = df[
        (df['zerer_kvartal'].isin(quarterly_dates)) &
        (df['odenis_kvartal'].isin(quarterly_dates))
        ]

    pivot_df = filtered_df.pivot_table(
        index='zerer_kvartal',
        columns='odenis_kvartal',
        values='VerilmisSigortaOdenisi',
        aggfunc='sum',
        fill_value=0
    ).reindex(index=quarterly_dates, columns=quarterly_dates, fill_value=0).cumsum(axis=1)

    pivot_df = shift_rows_left(pivot_df)

    pivot_df.columns = range(1, len(pivot_df.columns) + 1)

    # 13
    numeric_df = pivot_df.apply(pd.to_numeric, errors='coerce').fillna(0)
    col_sum = numeric_df.sum()
    pivot_df.loc[len(pivot_df) + 1] = col_sum

    # 14
    piv = filtered_df.pivot_table(
        columns='zerer_kvartal',
        values='VerilmisSigortaOdenisi',
        aggfunc='sum',
        fill_value=0
    ).reindex(columns=quarterly_dates, fill_value=0).reset_index(drop=True)

    pivot_row_13 = pivot_df.iloc[-1].tolist()
    piv_row = piv.iloc[0].tolist()
    result = [x - y for x, y in zip(pivot_row_13, piv_row)]
    result[-1] = 'X'
    pivot_df.loc[len(pivot_df) + 1] = result

    # 15
    coef_row = []
    for i in range(len(pivot_df.columns) - 1):
        numerator = pivot_df.iloc[-2, i + 1]
        denominator = pivot_df.iloc[-1, i]
        if denominator != 0 and denominator != 'X' and numerator != 'X':
            coef_row.append(numerator / denominator)
        else:
            coef_row.append('X')
    coef_row.append('X')
    pivot_df.loc[len(pivot_df) + 1] = coef_row

    # 16
    ink_amil = pivot_df.iloc[-1][:-1].tolist()
    result_list = []
    for i in range(len(ink_amil)):
        product = 1
        for j in range(i, len(ink_amil)):
            product *= ink_amil[j]
        result_list.append(product)
    result_list.append(1)
    pivot_df.loc[len(pivot_df) + 1] = result_list

    # 17
    last_row = pivot_df.iloc[-1]
    gec_amil = []

    for value in last_row:
        if value != 0 and value != 'X':
            gec_amil.append(1 / value)
        else:
            gec_amil.append('X')

    pivot_df.loc[len(pivot_df) + 1] = gec_amil
    return pivot_df


def form11(data, input_date):
    quarterly_dates = generate_quarterly_range(input_date)

    filtered_df = data[data['SigortaHadisesininBasverdiyiTarix'].apply(last_day_of_quarter).isin(quarterly_dates) &
                       data['VerilmisOdenisTarixi'].isna() & data['VerilmisSigortaOdenisi'].isna()]
    filtered_df1 = data[(data['SigortaHadisesininBasverdiyiTarix'] <= quarterly_dates[0]) &
                        data['VerilmisOdenisTarixi'].isna() & data['VerilmisSigortaOdenisi'].isna()]

    sorted_df = filtered_df.sort_values('SigortaHadisesininBasverdiyiTarix').reset_index(drop=True)
    unique_quarters = sorted(sorted_df['SigortaHadisesininBasverdiyiTarix'].apply(last_day_of_quarter).unique())
    new_df = pd.DataFrame(index=[''] + unique_quarters)
    new_df['Sətrin kodu'] = [f'{i:02d}' for i in range(1, len(new_df) + 1)]
    new_df['Bildirilmiş, lakin tənzimlənməmiş zərərlərin hesabat tarixinə məbləği'] = sorted_df.groupby(
        sorted_df['SigortaHadisesininBasverdiyiTarix'].apply(last_day_of_quarter))[
        'TekrarsigortacininBorcPayi'].sum().reindex(new_df.index, fill_value=0)
    new_df.iloc[0, 1] += filtered_df1['TekrarsigortacininBorcPayi'].sum()
    new_df['Bildirilmiş, lakin tənzimlənməmiş zərərlərdə təkrarsığortalının payı'] = 0
    new_df['Təkrarsığorta müqavilələrinə vaxtından əvvəl xitam verilməsi məbləği'] = (new_df.iloc[:, 1] + new_df.iloc[:,
                                                                                                          2]) * 0.03
    new_df['Bildirilmiş zərərlər ehtiyatında təkrarsığortaçıların payı'] = new_df.iloc[:, 1:4].sum(axis=1)
    new_df.loc['Yekun BTZE'] = ['X'] + list(new_df.iloc[:, 1:].sum())

    num_quarters = len(new_df.iloc[1:-2])
    quarter_names = [f'rüb {i}' for i in range(num_quarters, 0, -1)] + ['Hesabat tarixi ilə qurtaran rüb']
    new_df.index = [''] + quarter_names + ['Yekun BTZE']
    return new_df.reset_index().rename(columns={'index': 'Sığorta hadisələrinin baş verdiyi rüblər'})


form_3 = form3(dz, input_date)
form_8 = form8(dz, input_date)
form_11 = form11(dz, input_date)