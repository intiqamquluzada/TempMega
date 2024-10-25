import pandas as pd
from datetime import datetime

df = pd.read_excel('/Users/quluzade/Desktop/mega-prod/data_files/Müqavilələr jurnalı (satış növü)-.xlsx')


def tarix_formatla(date):
    aylar = [
        "yanvar", "fevral", "mart", "aprel", "may", "iyun",
        "iyul", "avqust", "sentyabr", "oktyabr", "noyabr", "dekabr"
    ]
    gun = date.day
    ay = aylar[date.month - 1]
    il = date.year
    return f"«{gun}» {ay} {il}"


input_date_str = input("Tarixi YYYY-MM-DD formatında daxil edin: ")
input_date = datetime.strptime(input_date_str, "%Y-%m-%d")
formatted_date = tarix_formatla(input_date)


# Tarixdən öncə xitam verilmiş müqavilələri gətirən funksiya
def active_date(data, date):
    filtered_data = data[data['SigortaTeminatininSonTarixi'] > date]
    return filtered_data


# Xitam verilmiş müqavilələri gətir
df1 = active_date(df, input_date)


def prepare_insurance(data, sales_type, invert=False):
    if invert:
        filtered_data = data[(data['Satış Növü'] != sales_type) & pd.isna(data['XitamVerildiyiTarix'])]
    else:
        filtered_data = data[(data['Satış Növü'] == sales_type) & pd.isna(data['XitamVerildiyiTarix'])]

    return filtered_data


def form1(data, columns):
    result_df = pd.DataFrame()
    result_df[columns[0]] = data['SigortaMuqavilesi']
    result_df[columns[1]] = data['SigortaMuqavilesiBaglandigiTarix']
    result_df[columns[2]] = data['Hesablanmisdir'].fillna(0)
    result_df[columns[3]] = (
        data['HesablananKomisyon'].where(
            data['HesablananKomisyon'] <= data['Hesablanmisdir'] * 0.15,
            data['Hesablanmisdir'] * 0.15
        )).fillna(0)
    result_df[columns[4]] = result_df[columns[2]] - result_df[columns[3]].fillna(0)
    result_df[columns[5]] = data['Hesablanmisdir_katastrofik'].fillna(0)
    result_df[columns[6]] = (
        data['HesablananKomisyon_katastrofik'].where(
            data['HesablananKomisyon_katastrofik'] <= data['Hesablanmisdir_katastrofik'] * 0.15,
            data['Hesablanmisdir_katastrofik'] * 0.15
        )).fillna(0)
    result_df[columns[7]] = result_df[columns[5]] - result_df[columns[6]].fillna(0)

    sums = result_df[columns[2:5]].sum()

    total_row = pd.Series(
        [None] * len(result_df.columns),
        index=result_df.columns
    )
    total_row[columns[0]] = 'Yekun BSH'
    total_row[columns[2:5]] = sums

    result_df = pd.concat([result_df, total_row.to_frame().T], ignore_index=True)

    return result_df


def form2(data, columns, input_date):
    result_df = pd.DataFrame()
    result_df[columns[0]] = data['SigortaMuqavilesi']
    result_df[columns[1]] = data['SigortaMeblegi'].fillna(0)
    result_df[columns[2]] = (
                data['SigortaTeminatininSonTarixi'] - data['SigortaTeminatininBaslangicTarixi']).dt.days.fillna(0)
    result_df[columns[3]] = (input_date - data['SigortaTeminatininBaslangicTarixi']).dt.days.fillna(0)
    result_df[columns[4]] = result_df[columns[1]] * (result_df[columns[2]] - result_df[columns[3]]) / result_df[
        columns[2]].fillna(0)
    result_df[columns[5]] = data['Hesablanmisdir_katastrofik'].fillna(0) - data['HesablananKomisyon_katastrofik'].where(
        data['HesablananKomisyon_katastrofik'] <= data['Hesablanmisdir_katastrofik'] * 0.15,
        data['Hesablanmisdir_katastrofik'] * 0.15
    ).fillna(0)
    result_df[columns[6]] = result_df[columns[5]] * (
            result_df[columns[2]] - result_df[columns[3]]) / result_df[columns[2]].fillna(0)

    sums = result_df.iloc[:, 4:7].sum()

    total_row = pd.Series(
        [None] * len(result_df.columns),
        index=result_df.columns
    )
    total_row[columns[0]] = 'Aralıq Yekun'
    total_row.iloc[4:7] = sums
    result_df = pd.concat([result_df, total_row.to_frame().T], ignore_index=True)

    return result_df


def form2yekun(data1, data2):
    column_sum_4 = data1.iloc[-1, 4] + data2.iloc[-1, 4]
    column_sum_6 = data1.iloc[-1, 6] + data2.iloc[-1, 6]
    yeni_setir = pd.DataFrame([['Yekun BSH'] + ['X'] + ['X'] + ['X'] + [column_sum_4] + ['X'] + [column_sum_6]],
                              columns=data2.columns)
    forma2_b = pd.concat([data2, yeni_setir], ignore_index=True)
    return forma2_b, column_sum_4, column_sum_6


def prepare_class(data, filter):
    result_df = data[data[filter] > 0]
    return result_df


def form4(data):
    columns = [
        "Təkrarsığorta müqavilələri (təkrarsığortaya  ötürülmüş risklər üzrə)",
        "Təkrarsığorta müqaviləsinin bağlandığı tarix",
        "Hesablanmış təkrarsığorta haqqı",
        "Komisyon muzdu",
        "Baza təkrarsığorta haqqı (III-IV)",
        "Hesablanmış təkrarsığorta haqqının katastrofik risk təminatına düşən hissəsi",
        "Katastrofik risk təminatı üzrə komisyon muzd",
        "Baza təkrarsığorta haqqının katastrofik risk təminatına düşən hissəsi  (VI-VII)"
    ]

    result_df = pd.DataFrame()
    result_df[columns[0]] = data['TekrarsigortaSlipininNömresi']
    result_df[columns[1]] = data['TekrarsigortaMuqavilesininBaglandigiTarix']
    result_df[columns[2]] = data["I_QrupTekrarsigortacilarPremiya"] + data["II_QrupTekrarsigortacilarPremiya"] + data[
        "III_QrupTekrarsigortacilarPremiya"] + data["DigerTekrarsigortacilarPremiya"]
    result_df[columns[3]] = data["I_QrupTekrarsigortacilarKomisyon"] + data["II_QrupTekrarsigortacilarKomisyon"] + data[
        "III_QrupTekrarsigortacilarKomisyon"] + data["DigerTekrarsigortacilarKomisyon"]
    result_df[columns[4]] = result_df[columns[2]] - result_df[columns[3]]
    result_df[columns[5]] = data["I_QrupTekrarsigortacilarPremiya_katastrofik"] + data[
        "II_QrupTekrarsigortacilarPremiya_katastrofik"] + data["III_QrupTekrarsigortacilarPremiya_katastrofik"] + data[
                                "DigerTekrarsigortacilarPremiya_katastrofik"]
    result_df[columns[6]] = data["I_QrupTekrarsigortacilarKomisyon_katastrofik"] + data[
        "II_QrupTekrarsigortacilarKomisyon_katastrofik"] + data["III_QrupTekrarsigortacilarKomisyon_katastrofik"] + \
                            data["DigerTekrarsigortacilarKomisyon_katastrofik"]
    result_df[columns[7]] = result_df[columns[5]] - result_df[columns[6]]
    sums = result_df[columns[4]].sum()

    total_row = pd.Series(
        [None] * len(result_df.columns),
        index=result_df.columns
    )
    total_row[columns[0]] = 'Aralıq Yekun'
    total_row[columns[4]] = sums

    result_df = pd.concat([result_df, total_row.to_frame().T], ignore_index=True)

    return result_df


def form5(data):
    columns = [
        "Təkrarsığorta müqavilələri (təkrarsığortaya verilmiş risklər üzrə)",
        "Baza təkrarsığorta haqqı",
        "Təkrarsığorta təminatının müddəti (günlərlə)",
        "Təkrarsığorta təminatının başlandığı andan hesabat tarixinə qədər günlərin sayı",
        "Qazanılmamış təkrarsığorta haqqı (IIx(III- IV)/III)",
        "Baza təkrarsığorta haqqının katastrofik risk təminatına düşən hissəsi",
        "Qazanılmamış təkrarsığorta haqqının katastrofik risk təminatına düşən hissəsi (VIx(III- IV)/III"
    ]

    result_df = pd.DataFrame()
    result_df[columns[0]] = data['TekrarsigortaSlipininNömresi']
    result_df[columns[1]] = (data["I_QrupTekrarsigortacilarPremiya"] + data["II_QrupTekrarsigortacilarPremiya"] + data[
        "III_QrupTekrarsigortacilarPremiya"] + data["DigerTekrarsigortacilarPremiya"]) - (
                                        data["I_QrupTekrarsigortacilarKomisyon"] + data[
                                    "II_QrupTekrarsigortacilarKomisyon"] + data["III_QrupTekrarsigortacilarKomisyon"] +
                                        data["DigerTekrarsigortacilarKomisyon"])
    result_df[columns[2]] = (
                data["TekrarsigortaTeminatininSonTarixi"] - data["TekrarsigortaTeminatininBaslangicTarixi"]).dt.days
    result_df[columns[3]] = (input_date - data["TekrarsigortaTeminatininBaslangicTarixi"]).dt.days
    result_df[columns[4]] = result_df[columns[1]] * (result_df[columns[2]] - result_df[columns[3]]) / result_df[
        columns[2]]
    result_df[columns[5]] = (
                                    data["I_QrupTekrarsigortacilarPremiya_katastrofik"] + data[
                                "II_QrupTekrarsigortacilarPremiya_katastrofik"] + data[
                                        "III_QrupTekrarsigortacilarPremiya_katastrofik"] + data[
                                        "DigerTekrarsigortacilarPremiya_katastrofik"]) - (
                                    data["I_QrupTekrarsigortacilarKomisyon_katastrofik"] + data[
                                "II_QrupTekrarsigortacilarKomisyon_katastrofik"] + data[
                                        "III_QrupTekrarsigortacilarKomisyon_katastrofik"] + data[
                                        "DigerTekrarsigortacilarKomisyon_katastrofik"])
    result_df[columns[6]] = result_df[columns[5]] * (result_df[columns[2]] - result_df[columns[3]]) / result_df[
        columns[2]]

    sums = result_df[columns[4]].sum()

    total_row = pd.Series(
        [None] * len(result_df.columns),
        index=result_df.columns
    )
    total_row[columns[0]] = 'Aralıq Yekun'
    total_row[columns[4]] = sums

    result_df = pd.concat([result_df, total_row.to_frame().T], ignore_index=True)

    return result_df


def form6(form_5):
    data = {
        "Təkrarsığortaçı-ların qrupları": [
            "I qrup təkrarsığortaçılar",
            "II qrup təkrarsığortaçılar",
            "III qrup təkrarsığortaçılar",
            "IV qrup təkrarsığortaçılar"
        ],
        "Qazanılmamış sığorta haqları ehtiyatının baza hissəsində təkrarsığortaçıların payı":
            [form.iloc[:, 4].sum().round() for form in form_5],
        "Təkrarsığortaçıların qrupları üzrə əmsallar, %": [0, 15, 25, 50]
    }

    df = pd.DataFrame(data)
    df["Qazanılmamış sığorta haqları ehtiyatının əlavə hissəsi (IIx III)"] = (
            df["Qazanılmamış sığorta haqları ehtiyatının baza hissəsində təkrarsığortaçıların payı"] *
            df["Təkrarsığortaçıların qrupları üzrə əmsallar, %"] / 100
    )

    total_row = pd.Series({
        "Təkrarsığortaçı-ların qrupları": "Yekun QSHEƏ",
        "Qazanılmamış sığorta haqları ehtiyatının baza hissəsində təkrarsığortaçıların payı": df.iloc[:, 1].sum(),
        "Təkrarsığortaçıların qrupları üzrə əmsallar, %": "X",
        "Qazanılmamış sığorta haqları ehtiyatının əlavə hissəsi (IIx III)": df.iloc[:, 3].sum()
    })

    return pd.concat([df, total_row.to_frame().T], ignore_index=True)


columns_form1_a = [
    'Sığorta müqavilələri',
    'Sığorta müqaviləsinin bağlandığı tarix',
    'Hesablanmış sığorta haqqı',
    'Hesablanmış komisyon muzd (hesablanmış sığorta haqqının 15%-dən çox olmamaqla)',
    'Baza sığorta haqqı (III-IV)',
    'Hesablanmış sığorta haqqının katastrofik risk təminatına düşən hissəsi',
    'Katastrofik risk təminatı üzrə komisyon muzd (hesablanmış sığorta haqqının katastrofik risk təminatına düşən hissəsinin 15%-dən çox olmamaqla)',
    'Baza sığorta haqqının katastrofik risk təminatına düşən hissəsi (VI-VII)'
]

columns_form1_b = [
    'Təkrarsığorta müqavilələri (təkrarsığortaya qəbul edilmiş risklər üzrə)',
    'Təkrarsığorta müqaviləsinin bağlandığı tarix',
    'Hesablanmış təkrarsığorta haqqı',
    'Hesablanmış komisyon muzd (hesablanmış sığorta haqqının 15%-dən çox olmamaqla)',
    'Baza təkrarsığorta haqqı (III-IV)',
    'Hesablanmış təkrarsığorta haqqının katastrofik risk təminatına düşən hissəsi',
    'Katastrofik risk təminatı üzrə komisyon muzd (hesablanmış sığorta haqqının katastrofik risk təminatına düşən hissəsinin 15%-dən çox olmamaqla)',
    'Baza sığorta haqqının katastrofik risk təminatına düşən hissəsi (VI-VII)'
]

columns_form2_a = [
    'Sığorta müqavilələri',
    'Baza sığorta haqqı',
    'Sığorta təminatının müddəti (günlərlə)',
    'Sığorta təminatının başlandığı andan hesabat tarixinə qədər günlərin sayı',
    'Qazanılmamış sığorta haqqı (IIx(III-IV)/III)',
    'Baza sığorta haqqının katastrofik risk təminatına düşən hissəsi',
    'Qazanılmamış sığorta haqqının katastrofik risk təminatına düşən hissəsi (VIx(III-IV)/III)'
]

columns_form2_b = [
    'Təkrarsığorta müqavilələri (təkrarsığortaya qəbul edilmiş risklər üzrə)',
    'Baza təkrarsığorta haqqı',
    'Təkrarsığorta təminatının müddəti (günlərlə)',
    'Təkrarsığorta təminatının başlandığı andan hesabat tarixinə qədər günlərin sayı',
    'Qazanılmamış təkrarsığorta haqqı (IIx(III- IV)/III)',
    'Baza təkrarsığorta haqqının katastrofik risk təminatına düşən hissəsi',
    'Qazanılmamış təkrarsığorta haqqının katastrofik risk təminatına düşən hissəsi (VIx(III-IV)/III)'
]

insurance_groups = [
    "I_QrupTekrarsigortacilarPremiya",
    "II_QrupTekrarsigortacilarPremiya",
    "III_QrupTekrarsigortacilarPremiya",
    "DigerTekrarsigortacilarPremiya"
]

filtered_a = prepare_insurance(df1, 'ALINAN-FAKULTATİV TƏKRAR SIĞORTA', invert=True)
filtered_b = prepare_insurance(df1, 'ALINAN-FAKULTATİV TƏKRAR SIĞORTA')
forma1_a = form1(filtered_a, columns_form1_a)
forma1_b = form1(filtered_b, columns_form1_b)
forma2_a = form2(filtered_a, columns_form2_a, input_date)
forma2_b = form2(filtered_b, columns_form2_b, input_date)
forma2_b = form2yekun(forma2_a, forma2_b)
form_4 = [form4(prepare_class(df1, grp)) for grp in insurance_groups]
form_5 = [form5(prepare_class(df1, grp)) for grp in insurance_groups]
form_6 = form6(form_5)