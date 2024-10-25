import pandas as pd
from datetime import datetime


# df = pd.read_excel('C:/Users/Hp/Desktop/mega/Müqavilələr jurnalı (satış növü)-.xlsx')
# dz = pd.read_excel('C:/Users/Hp/Desktop/mega/ZJ 5 illik.xlsx', engine='openpyxl')
# dfs = pd.read_excel(r'C:\Users\Hp\Desktop\mega\Subroqasiya__sas.xlsx')
# db = pd.read_excel('C:/Users/Hp/Desktop/mega/Forma 8_7.xlsx')
# # dq = pd.read_excel(r'C:\Users\Hp\Desktop\testing\Avtokasko.xlsx',sheet_name="Forma8_10")
# input_date_str = input("Tarixi YYYY-MM-DD formatında daxil edin: ")
# # input_date = datetime.strptime(input_date_str, "%Y-%m-%d")


def total(df, dz, dfs, db, dq, input_date_str, sinif, year):
    input_date = datetime.strptime(input_date_str, "%Y-%m-%d")

    def tarix_formatla(date):
        aylar = [
            "yanvar", "fevral", "mart", "aprel", "may", "iyun",
            "iyul", "avqust", "sentyabr", "oktyabr", "noyabr", "dekabr"
        ]
        gun = date.day
        ay = aylar[date.month - 1]
        il = date.year
        return f"«{gun}» {ay} {il}"

    # Tarixdən öncə xitam verilmiş müqavilələri gətirən funksiya
    def active_date(data, data1, data2, sinif, date):

        new_column_names = data2.iloc[0, 5:8].tolist()
        data2.columns = list(data2.columns[:5]) + new_column_names + list(data2.columns[8:])
        data2 = data2.iloc[3:].reset_index(drop=True)
        data2['Subroqasiya gəlirinin daxil olduğu tarix'] = pd.to_datetime(
            data2['Subroqasiya gəlirinin daxil olduğu tarix'], errors='coerce')

        data2.columns = data2.columns.str.strip()
        df_filtered = data[(data['SigortaTeminatininSonTarixi'] < date) & (data['SigortaSinifi'] == sinif)]
        dz_filtered = data1[(data1['SigortaHadisesininBasverdiyiTarix'] < date) & (data1['SigortaSinfi'] == sinif)]
        ds_filtered = data2[(data2['Subroqasiya gəlirinin daxil olduğu tarix'] < date) & (data2['Sığоrtаnın sinifləri'] == sinif)]

        return df_filtered, dz_filtered, ds_filtered

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

        sums = result_df[columns[2:8]].sum()

        total_row = pd.Series(
            [None] * len(result_df.columns),
            index=result_df.columns
        )
        total_row[columns[0]] = 'Yekun BSH'
        total_row[columns[2:8]] = sums

        if not total_row.dropna().empty:
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
        result_df[columns[5]] = data['Hesablanmisdir_katastrofik'].fillna(0) - data[
            'HesablananKomisyon_katastrofik'].where(
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
        if not total_row.dropna().empty:
            result_df = pd.concat([result_df, total_row.to_frame().T], ignore_index=True)

        return result_df

    def form2yekun(data1, data2):
        column_sum_4 = data1.iloc[-1, 4] + data2.iloc[-1, 4]
        column_sum_5 = data1.iloc[-1, 5] + data2.iloc[-1, 5]
        column_sum_6 = data1.iloc[-1, 6] + data2.iloc[-1, 6]
        yeni_setir = pd.DataFrame(
            [['Yekun BSH'] + ['X'] + ['X'] + ['X'] + [column_sum_4] + [column_sum_5] + [column_sum_6]],
            columns=data2.columns)

        result_df = pd.concat([data2, yeni_setir], ignore_index=True)
        return result_df, column_sum_4, column_sum_6

    def prepare_class(data, filter):
        result_df = data[data[filter] > 0]
        return result_df

    def form4(data):
        columns = [
            "Təkrarsığorta müqavilələri (təkrarsığortaya  ötürülmüş risklər üzrə)",
            "Təkrarsığorta müqaviləsinin bağlandığı tarix",
            "Hesablanmış təkrarsığorta haqqı",
            "Komisyon muzdu",
            "Baza təkrarsığorta haqqı (III-IV)",
            "Hesablanmış təkrarsığorta haqqının katastrofik risk təminatına düşən hissəsi",
            "Katastrofik risk təminatı üzrə komisyon muzd",
            "Baza təkrarsığorta haqqının katastrofik risk təminatına düşən hissəsi  (VI-VII)"
        ]

        result_df = pd.DataFrame()
        result_df[columns[0]] = data['TekrarsigortaSlipininNömresi']
        result_df[columns[1]] = data['TekrarsigortaMuqavilesininBaglandigiTarix']
        result_df[columns[2]] = data["I_QrupTekrarsigortacilarPremiya"] + data["II_QrupTekrarsigortacilarPremiya"] + \
                                data[
                                    "III_QrupTekrarsigortacilarPremiya"] + data["DigerTekrarsigortacilarPremiya"]
        result_df[columns[3]] = data["I_QrupTekrarsigortacilarKomisyon"] + data["II_QrupTekrarsigortacilarKomisyon"] + \
                                data[
                                    "III_QrupTekrarsigortacilarKomisyon"] + data["DigerTekrarsigortacilarKomisyon"]
        result_df[columns[4]] = (result_df[columns[2]] - result_df[columns[3]]).fillna(0)
        result_df[columns[5]] = data["I_QrupTekrarsigortacilarPremiya_katastrofik"] + data[
            "II_QrupTekrarsigortacilarPremiya_katastrofik"] + data["III_QrupTekrarsigortacilarPremiya_katastrofik"] + \
                                data[
                                    "DigerTekrarsigortacilarPremiya_katastrofik"]
        result_df[columns[6]] = data["I_QrupTekrarsigortacilarKomisyon_katastrofik"] + data[
            "II_QrupTekrarsigortacilarKomisyon_katastrofik"] + data["III_QrupTekrarsigortacilarKomisyon_katastrofik"] + \
                                data["DigerTekrarsigortacilarKomisyon_katastrofik"]
        result_df[columns[7]] = (result_df[columns[5]] - result_df[columns[6]]).fillna(0)
        sums = result_df[columns[4:7]].sum()

        total_row = pd.Series(
            [None] * len(result_df.columns),
            index=result_df.columns
        )
        total_row[columns[0]] = 'Aralıq Yekun'
        total_row[columns[4:7]] = sums

        result_df = pd.concat([result_df, total_row.to_frame().T], ignore_index=True)

        return result_df

    def form4yekun():
        column_sum_4 = sum(form.iloc[-1, 4] for form in form_4 if pd.notna(form.iloc[-1, 4]))
        column_sum_5 = sum(form.iloc[-1, 5] for form in form_4 if pd.notna(form.iloc[-1, 5]))
        column_sum_6 = sum(form.iloc[-1, 6] for form in form_4 if pd.notna(form.iloc[-1, 6]))
        column_sum_7 = sum(form.iloc[-1, 7] for form in form_4 if pd.notna(form.iloc[-1, 7]))
        yeni_setir = pd.DataFrame(
            [['Yekun BSH'] + ['X'] + ['X'] + ['X'] + [column_sum_4] + [column_sum_5] + [column_sum_6] + [column_sum_7]],
            columns=form_4[0].columns)

        result_df = pd.concat([form_4[3], yeni_setir], ignore_index=True)
        return result_df, column_sum_4, column_sum_7

    def form5(data):
        columns = [
            "Təkrarsığorta müqavilələri (təkrarsığortaya verilmiş risklər üzrə)",
            "Baza təkrarsığorta haqqı",
            "Təkrarsığorta təminatının müddəti (günlərlə)",
            "Təkrarsığorta təminatının başlandığı andan hesabat tarixinə qədər günlərin sayı",
            "Qazanılmamış təkrarsığorta haqqı (IIx(III- IV)/III)",
            "Baza təkrarsığorta haqqının katastrofik risk təminatına düşən hissəsi",
            "Qazanılmamış təkrarsığorta haqqının katastrofik risk təminatına düşən hissəsi (VIx(III- IV)/III"
        ]

        result_df = pd.DataFrame()
        result_df[columns[0]] = data['TekrarsigortaSlipininNömresi']
        result_df[columns[1]] = (data["I_QrupTekrarsigortacilarPremiya"] + data["II_QrupTekrarsigortacilarPremiya"] +
                                 data[
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

        sums = result_df[columns[4:7]].sum()
        total_row = pd.Series(
            [None] * len(result_df.columns),
            index=result_df.columns
        )
        total_row[columns[0]] = 'Aralıq Yekun'
        total_row[columns[4:7]] = sums

        result_df = pd.concat([result_df, total_row.to_frame().T], ignore_index=True)
        return result_df

    def form5yekun():
        column_sum_4 = sum(form.iloc[-1, 4] for form in form_5 if pd.notna(form.iloc[-1, 4]))
        column_sum_5 = sum(form.iloc[-1, 5] for form in form_5 if pd.notna(form.iloc[-1, 5]))
        column_sum_6 = sum(form.iloc[-1, 6] for form in form_5 if pd.notna(form.iloc[-1, 6]))

        yeni_setir = pd.DataFrame(
            [['Yekun BSH'] + ['X'] + ['X'] + ['X'] + [column_sum_4] + [column_sum_5] + [column_sum_6]],
            columns=form_5[3].columns)

        result_df = pd.concat([form_5[3], yeni_setir], ignore_index=True)
        return result_df, column_sum_4, column_sum_6

    def form6():
        data = {
            "Təkrarsığortaçı-ların qrupları": [
                "I qrup təkrarsığortaçılar",
                "II qrup təkrarsığortaçılar",
                "III qrup təkrarsığortaçılar",
                "IV qrup təkrarsığortaçılar"
            ],
            "Qazanılmamış sığorta haqları ehtiyatının baza hissəsində təkrarsığortaçıların payı":
                (form.iloc[-1, 4] for form in form_5 if pd.notna(form.iloc[-1, 4])),
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

    def form7(data):
        filtered_df = data.iloc[:, 3:11]
        filtered_df = filtered_df[filtered_df[filtered_df.columns[0]].notna()]
        filtered_df = filtered_df.reset_index(drop=True)
        column_names = filtered_df.iloc[1].tolist()
        filtered_df.columns = column_names
        filtered_df = filtered_df.drop([0, 1, 2]).reset_index(drop=True)
        # filtered_df = filtered_df.shift(-1)
        filtered_df.iloc[0, 0:8] = ['III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X']
        quarter_names = [f'rüb {i + 1}' for i in range(len(filtered_df) - 1)]
        data2 = pd.DataFrame({
            'Zərərlərin baş verdiyi rüblər': ['I'] + quarter_names + ['Hesabat Tarixi ile Qurtaran rub'],
            'Sətrin kodu': ['II'] + [f'{i + 1:02d}' for i in range(len(filtered_df) - 1)] + ['']
        })
        data2 = data2.reset_index(drop=True)
        filtered_df = filtered_df.reset_index(drop=True).fillna(0)
        filt = filtered_df.iloc[-1, 2]
        filt1 = filtered_df.iloc[-1, 6]
        result_df = pd.concat([data2, filtered_df], axis=1)
        result_df.iloc[-1, [2, 3, 4, 5, 6, 7, 8, 9]] = [form7_1a[0], filt, form_2yekun[1],
                                                        form7_1a[0] + filt - form_2yekun[1],
                                                        form7_1a[1], filt1, form_2yekun[2],
                                                        form7_1a[1] + filt1 - form_2yekun[2]]

        return result_df

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

    def end_quarter(data):
        data['Quarter'] = data['SigortaMuqavilesiBaglandigiTarix'].apply(last_day_of_quarter)
        data = data[data["Quarter"] == data["Quarter"].max()]
        data["Bazasığorta haqqı"] = data['Hesablanmisdir'] - data['HesablananKomisyon'].where(
            data['HesablananKomisyon'] <= data['Hesablanmisdir'] * 0.15,
            data['Hesablanmisdir'] * 0.15
        )
        data['Katastorfik Bazasığorta haqqı'] = data['Hesablanmisdir_katastrofik'] - data[
            'HesablananKomisyon_katastrofik'].where(
            data['HesablananKomisyon_katastrofik'] <= data['Hesablanmisdir_katastrofik'] * 0.15,
            data['Hesablanmisdir_katastrofik'] * 0.15
        )
        sum1 = data["Bazasığorta haqqı"].sum()
        sum2 = data['Katastorfik Bazasığorta haqqı'].sum()
        return sum1, sum2

    def end_quarter10(data):
        data['Quarter'] = data['TekrarsigortaMuqavilesininBaglandigiTarix'].apply(last_day_of_quarter)
        data = data[data["Quarter"] == data["Quarter"].max()]
        data["Təkrar Bazasığorta haqqı"] = (data['I_QrupTekrarsigortacilarPremiya_katastrofik'] + data[
            'II_QrupTekrarsigortacilarPremiya_katastrofik'] +
                                            data['III_QrupTekrarsigortacilarPremiya_katastrofik'] + data[
                                                'DigerTekrarsigortacilarPremiya_katastrofik'])

        data['Katastrofik Təkrarsiğorta haqqı'] = (data['I_QrupTekrarsigortacilarPremiya_katastrofik'] + data[
            'II_QrupTekrarsigortacilarPremiya_katastrofik'] +
                                                   data['III_QrupTekrarsigortacilarPremiya_katastrofik'] + data[
                                                       'DigerTekrarsigortacilarPremiya_katastrofik']) - (
                                                          data['I_QrupTekrarsigortacilarKomisyon_katastrofik'] + data[
                                                      'II_QrupTekrarsigortacilarKomisyon_katastrofik'] +
                                                          data['III_QrupTekrarsigortacilarKomisyon_katastrofik'] + data[
                                                              'DigerTekrarsigortacilarKomisyon_katastrofik'])
        sum1 = data["Təkrar Bazasığorta haqqı"].sum()
        sum2 = data['Katastrofik Təkrarsiğorta haqqı'].sum()
        return sum1, sum2

    def form10(data):
        filtered_df = data.iloc[:, 3:11]
        filtered_df = filtered_df[filtered_df[filtered_df.columns[0]].notna()]
        filtered_df = filtered_df.reset_index(drop=True)
        column_names = filtered_df.iloc[1].tolist()
        filtered_df.columns = column_names
        filtered_df = filtered_df.drop([0, 1, 2]).reset_index(drop=True)
        # filtered_df = filtered_df.shift(-1)
        filtered_df.iloc[0, 0:8] = ['III', 'IV', 'V', 'VI', 'VII', 'VIII', 'IX', 'X']
        quarter_names = [f'rüb {i + 1}' for i in range(len(filtered_df) - 1)]
        data2 = pd.DataFrame({
            'Zərərlərin baş verdiyi rüblər': ['I'] + quarter_names + ['Hesabat Tarixi ile Qurtaran rub'],
            'Sətrin kodu': ['II'] + [f'{i + 1:02d}' for i in range(len(filtered_df))]
        })
        data2 = data2.reset_index(drop=True)
        filtered_df = filtered_df.reset_index(drop=True).fillna(0)
        filt = filtered_df.iloc[-1, 2]
        filt1 = filtered_df.iloc[-1, 6]
        result_df = pd.concat([data2, filtered_df], axis=1)
        result_df.iloc[-1, [2, 3, 4, 5, 6, 7, 8, 9]] = [form10_4a[0], filt, form_5yekun[1],
                                                        form10_4a[0] + filt - form_5yekun[1]
            , form10_4a[1], filt1, form_5yekun[2], form10_4a[1] + filt1 - form_5yekun[2]]

        return result_df

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

    def generate_quarterly_range(input_date, year_type):
        end_date = last_day_of_quarter(input_date)
        start_date = end_date - pd.DateOffset(years=year_type) + pd.DateOffset(days=3)

        quarters = pd.date_range(start=start_date, end=end_date, freq='Q')
        quarter_end_dates = quarters.map(last_day_of_quarter)

        return quarter_end_dates

    def form3(data):
        quarterly_dates = generate_quarterly_range(input_date,year)
        
        filtered_df = data[
            data['SigortaHadisesininBasverdiyiTarix'].apply(last_day_of_quarter).isin(quarterly_dates) &
            data['VerilmisOdenisTarixi'].isna() & 
            data['VerilmisSigortaOdenisi'].isna()
        ]
        
        sorted_df = filtered_df.sort_values('SigortaHadisesininBasverdiyiTarix').reset_index(drop=True)    
        all_quarters = pd.DataFrame(index=[''] + list(quarterly_dates))
        all_quarters['Sətrin kodu'] = [f'{i:02d}' for i in range(1, len(all_quarters) + 1)]
        all_quarters['Bildirilmiş, lakin tənzimlənməmiş zərərlərin hesabat tarixinə məbləği'] = (
            sorted_df.groupby(sorted_df['SigortaHadisesininBasverdiyiTarix'].apply(last_day_of_quarter))[
                'SigortaOdenisiUzreCemiBorc'
            ].sum().reindex(all_quarters.index, fill_value=0)
        )

        all_quarters.iloc[0, 1] += data[
            (data['SigortaHadisesininBasverdiyiTarix'] <= quarterly_dates[0]) &
            data['VerilmisOdenisTarixi'].isna() & 
            data['VerilmisSigortaOdenisi'].isna()
        ]['SigortaOdenisiUzreCemiBorc'].sum()
        all_quarters['Sığorta müqavilələrinə vaxtından əvvəl xitam verilməsi məbləği'] = 0
        all_quarters['Zərərlərin tənzimlənməsi xərcləri (III+ IV)x0,03'] = all_quarters.iloc[:, 1] * 0.03
        all_quarters['Bildirilmiş zərərlər ehtiyatı (III+IV+V)'] = all_quarters.iloc[:, 1:4].sum(axis=1)
        all_quarters.loc['Yekun BTZE'] = ['X'] + list(all_quarters.iloc[:, 1:].sum())
        all_quarters.index = [''] + [f'rüb {i}' for i in range(len(quarterly_dates), 0, -1)] + ['Yekun BTZE']
        
        return all_quarters.reset_index().rename(columns={'index': 'Sığorta hadisələrinin baş verdiyi rüblər'})

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

    def is_float(value):
        try:
            return float(value)
        except ValueError:
            return 0

    def shift_x(df):
        for i in range(1, len(df)):
            if i <= len(df.columns):
                df.iloc[:, -i:] = df.iloc[:, -i:].astype('object')
                df.iloc[i, -i:] = 'X'
            else:
                df.iloc[i, :] = 'X'
        return df

    ###Bunlara Sinif girecek
    def emsal(class_of_insurance):
        insurance_classes = [
            '(01)FerdiQeza', '(01)FerdiQeza',
            '(02)Tibbi', '(02)Tibbi',
            '(03)EmlakYanginDigerRisk', '(03)EmlakYanginDigerRisk',
            '(04)AvtoKasko', '(04)AvtoKasko',
            '(05)DemiryolNeqliyyVasitesi', '(05)DemiryolNeqliyyVasitesi',
            '(06)HavaNeqliyyKasko', '(06)HavaNeqliyyKasko',
            '(07)SuNeqliyyKasko', '(07)SuNeqliyyKasko',
            '(08)Yuk', '(08)Yuk',
            '(09)KendTeserrufBitki', '(09)KendTeserrufBitki',
            '(10)KendTeserrufHeyvan', '(10)KendTeserrufHeyvan',
            '(11)IshcilerinDeleduzlug', '(11)IshcilerinDeleduzlug',
            '(12)PulvePulSenedSaxtalash', '(12)PulvePulSenedSaxtalash'
        ]
        forms = [
            'Forma 8_8 üzrə', 'Forma 8_12 üzrə',
            'Forma 8_8 üzrə', 'Forma 8_12 üzrə',
            'Forma 8_8 üzrə', 'Forma 8_12 üzrə',
            'Forma 8_8 üzrə', 'Forma 8_12 üzrə',
            'Forma 8_8 üzrə', 'Forma 8_12 üzrə',
            'Forma 8_8 üzrə', 'Forma 8_12 üzrə',
            'Forma 8_8 üzrə', 'Forma 8_12 üzrə',
            'Forma 8_8 üzrə', 'Forma 8_12 üzrə',
            'Forma 8_8 üzrə', 'Forma 8_12 üzrə',
            'Forma 8_8 üzrə', 'Forma 8_12 üzrə',
            'Forma 8_8 üzrə', 'Forma 8_12 üzrə',
            'Forma 8_8 üzrə', 'Forma 8_12 üzrə'
        ]
        emsal = [
            [1.5240, 1.0890, 1.0489, 1.0130, 1.0078, 1.0030, 1.0144, 1.0022, 1.0003, 1.0000, 1.0000, 1],
            [5.4074, 1.0324, 1.0263, 1.0026, 1.0014, 1.0045, 1.0045, 1.0029, 1.0000, 1.0000, 1.0000, 1],
            [1.3733, 1.0785, 1.0122, 1.0025, 1.0009, 1.0008, 1.0006, 1.0005, 1.0004, 1.0003, 1.0002, 1],
            [1.2076, 1.1993, 1.0117, 1.0012, 1.0012, 1.0001, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1],
            [1.3565, 1.1048, 1.0391, 1.0364, 1.0351, 1.0138, 1.0029, 1.0014, 1.0004, 1.0000, 1.0000, 1],
            [1.4542, 1.1029, 1.0661, 1.0640, 1.0624, 1.0268, 1.0025, 1.0000, 1.0000, 1.0000, 1.0000, 1],
            [1.8928, 1.0838, 1.0192, 1.0050, 1.0044, 1.0021, 1.0013, 1.0009, 1.0007, 1.0005, 1.0003, 1],
            [1.5959, 1.1003, 1.0456, 1.0311, 1.0142, 1.0141, 1.0140, 1.0134, 1.0026, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [1.9855, 1.5715, 1.0282, 1.2725, 1.0035, 1.0003, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1],
            [2.0587, 1.5884, 1.0119, 1.2974, 1.0029, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [3.4011, 1.6962, 1.0451, 1.0158, 1.0105, 1.0199, 0.9999, 1.0699, 0.9999, 0.9997, 1.0000, 1],
            [1.5537, 1.0992, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1],
            [1.5226, 1.0231, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1],
            [1.7017, 1.1620, 1.0408, 1.0210, 1.0180, 1.0086, 1.0086, 1.0000, 1.0000, 1.0000, 1.0000, 1],
            [1.7017, 1.1620, 1.0408, 1.0210, 1.0180, 1.0086, 1.0086, 1.0000, 1.0000, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1]
        ]
        insurance_classes1 = [
            "(13)AvtoKonulluMesuliyy",
            "(13)AvtoKonulluMesuliyy",
            "(14)DemiryolNeqliySahibMesuliyy",
            "(14)DemiryolNeqliySahibMesuliyy",
            "(15)HavaNeqliySahibMesuliyy",
            "(15)HavaNeqliySahibMesuliyy",
            "(16)SuNeqliySahibMesuliyy",
            "(16)SuNeqliySahibMesuliyy",
            "(17)YukDashiyanMesuliyy",
            "(17)YukDashiyanMesuliyy",
            "(18)MulkiMuqavileUzreMesuliyy",
            "(18)MulkiMuqavileUzreMesuliyy",
            "(19)PesheMesuliyy",
            "(19)PesheMesuliyy",
            "(20)IshegoturenMesuliyy",
            "(20)IshegoturenMesuliyy",
            "(21)UmumiMulkiMesuliyy",
            "(21)UmumiMulkiMesuliyy",
            "(22)Kredit",
            "(22)Kredit",
            "(23)Ipoteka",
            "(23)Ipoteka",
            "(24)EmlakinDeyerdenDushmesi",
            "(24)EmlakinDeyerdenDushmesi",
            "(25)IshinDayanmasiRiski",
            "(25)IshinDayanmasiRiski",
            "(30)DeputatlarinIcbari",
            "(30)DeputatlarinIcbari",
            "(31)TibbiPersonalinAIDSden",
            "(31)TibbiPersonalinAIDSden",
            "(32)HerbiQulluqcularinIcbari",
            "(32)HerbiQulluqcularinIcbari",
            "(33)HuquqMuhafizeIcbari",
            "(33)HuquqMuhafizeIcbari",
            "(34)DovletQulluqcuIcbari",
            "(34)DovletQulluqcuIcbari",
            "(35)DiplomatlarinIcbari",
            "(35)DiplomatlarinIcbari",
            "(36)AuditorPesheMesuliyyIcbari",
            "(36)AuditorPesheMesuliyyIcbari",
            "(37)IcbariDashinmazEmlak",
            "(37)IcbariDashinmazEmlak",
            "(38)IcbariDashinmazEmlakMesul",
            "(38)IcbariDashinmazEmlakMesul",
            "(39)IcbariNVSMMS",
            "(39)IcbariNVSMMS",
            "(40)IcbariSernishinFerdiQeza",
            "(40)IcbariSernishinFerdiQeza",
            "(41)Sefer",
            "(41)Sefer",
            "titul sığortası",
            "titul sığortası",
            "hüquqi xərclərin sığortası",
            "hüquqi xərclərin sığortası",
            "həyatın ölüm halından sığortası",
            "həyatın ölüm halından sığortası",
            "İstehsalatda bədbəxt hadisələr və peşə xəstəlikləri nəticəsində peşə əmək qabiliyyətinin itirilməsi hallarından icbari sığorta",
            "İstehsalatda bədbəxt hadisələr və peşə xəstəlikləri nəticəsində peşə əmək qabiliyyətinin itirilməsi hallarından icbari sığorta"
        ]
        forms1 = [
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 8_8 üzrə",
            "Forma 8_12 üzrə",
            "Forma 9_11 üzrə",
            "Forma 9_15 üzrə",
            "Forma 9_11 üzrə",
            "Forma 9_15 üzrə"
        ]
        emsal1 = [
            [1.5028, 1.1192, 1.0172, 1.0076, 1.0068, 1.0032, 1.0032, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [1.4488, 1.1533, 1.0679, 1.0003, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [3.1683, 1.0573, 1.0407, 1.0334, 1.0156, 1.0164, 1.0025, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [3.1683, 1.0573, 1.0407, 1.0334, 1.0156, 1.0164, 1.0025, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [3.1683, 1.0573, 1.0407, 1.0334, 1.0156, 1.0164, 1.0025, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [3.1683, 1.0573, 1.0407, 1.0334, 1.0156, 1.0164, 1.0025, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [3.1683, 1.0573, 1.0407, 1.0334, 1.0156, 1.0164, 1.0025, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [3.1683, 1.0573, 1.0407, 1.0334, 1.0156, 1.0164, 1.0025, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [3.1683, 1.0573, 1.0407, 1.0334, 1.0156, 1.0164, 1.0025, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [3.1683, 1.0573, 1.0407, 1.0334, 1.0156, 1.0164, 1.0025, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [3.1683, 1.0573, 1.0407, 1.0334, 1.0156, 1.0164, 1.0025, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [3.1683, 1.0573, 1.0407, 1.0334, 1.0156, 1.0164, 1.0025, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [3.1683, 1.0573, 1.0407, 1.0334, 1.0156, 1.0164, 1.0025, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [3.1683, 1.0573, 1.0407, 1.0334, 1.0156, 1.0164, 1.0025, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [3.1683, 1.0573, 1.0407, 1.0334, 1.0156, 1.0164, 1.0025, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [3.1683, 1.0573, 1.0407, 1.0334, 1.0156, 1.0164, 1.0025, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [3.1683, 1.0573, 1.0407, 1.0334, 1.0156, 1.0164, 1.0025, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [3.1683, 1.0573, 1.0407, 1.0334, 1.0156, 1.0164, 1.0025, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [1.3158, 1.0334, 1.0002, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [1.3158, 1.0334, 1.0002, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [3.1683, 1.0573, 1.0407, 1.0334, 1.0156, 1.0164, 1.0025, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [3.1683, 1.0573, 1.0407, 1.0334, 1.0156, 1.0164, 1.0025, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [1.8630, 1.0925, 1.0746, 1.0158, 1.0100, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1],
            [1.3083, 1.0248, 3.3027, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1],
            [3.1683, 1.0573, 1.0407, 1.0334, 1.0156, 1.0164, 1.0025, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [3.1683, 1.0573, 1.0407, 1.0334, 1.0156, 1.0164, 1.0025, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [1.6582, 1.0763, 1.0213, 1.0102, 1.0017, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [1.6582, 1.0763, 1.0213, 1.0102, 1.0017, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [2.3527, 1.4751, 1.0317, 1.0174, 1.0106, 1.0153, 1.0014, 1.0525, 1.0001, 1.0000, 1.0000, 1],
            [3.1683, 1.0573, 1.0407, 1.0334, 1.0156, 1.0164, 1.0025, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [3.1683, 1.0573, 1.0407, 1.0334, 1.0156, 1.0164, 1.0025, 1.0000, 1.0000, 1.0000, 1.0000, 1, 1, 1, 1, 1, 1,
             1, 1,
             1],
            [2.6284, 1.1534, 1.0113, 1.0010, 1.0008, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1],
            [2.6284, 1.1066, 1.0000, 1.1260, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1.0000, 1],
            [2.6496, 1.3038, 1.3037, 1.7259, 1.0391, 1.0305, 1.0158, 1.0143, 1.0027, 1.0000, 1.0000, 1],
            [3.3399, 1.3968, 1.2333, 2.7365, 1.0032, 1.0890, 1.0160, 1.0587, 1.0000, 1.0000, 1.0000, 1]
        ]

        emsals = pd.DataFrame({
            "SigortaSinifi": insurance_classes,
            "Forms": forms,
            "Emsal": emsal
        }
        )
        emsals1 = pd.DataFrame({
            "SigortaSinifi": insurance_classes1,
            "Forms": forms1,
            "Emsal": emsal1
        }
        )

        emsal_df = pd.concat([emsals, emsals1], ignore_index=True)

        evez_list8 = emsal_df[(emsal_df['SigortaSinifi'] == f'{class_of_insurance}') & (
                emsal_df['Forms'] == 'Forma 8_8 üzrə')].iloc[:, 2:].values.tolist()[0]
        evez_list12 = emsal_df[(emsal_df['SigortaSinifi'] == f'{class_of_insurance}') & (
                emsal_df['Forms'] == 'Forma 8_12 üzrə')].iloc[:, 2:].values.tolist()[0]

        return evez_list8, evez_list12

    def avgemsal():
        avg = {'SıgortaSinifi': ['(01)FerdiQeza', '(02)Tibbi', '(03)EmlakYanginDigerRisk', '(04)AvtoKasko',
                                 '(05)Demiryol Neqliyy Vasitesi', '(06)HavaNeqliyyKasko', '(07)SuNeqliyyKasko',
                                 '(08)Yuk',
                                 '(09)Kend Teserruf Bitki', '(10)KendTeserruf Heyvan', '(11)IshcilerinDeleduzlug',
                                 '(12)PulvePulSened Saxtalash', '(13)AvtoKonulluMesuliyy',
                                 '(14)Demiryol Neqliy SahibMesuliyy', '(15)HavaNeqliy SahibMesuliyy',
                                 '(16)SuNeqliy Sahib Mesuliyy', '(17)YukDashiyanMesuliyy',
                                 '(18)MulkiMuqavileUzreMesuliyy',
                                 '(19)PesheMesuliyy', '(20)Ishegoturen Mesuliyy', '(21)Umumi Mulki Mesuliyy',
                                 '(22)Kredit',
                                 '(23)Ipoteka', '(24)EmlakinDeyerdenDushmesi', '(25)IshinDayanmasiRiski',
                                 '(30)DeputatlarinIcbari', '(31)Tibbi Personalin AIDSden',
                                 '(32)HerbiQulluqcularinIcbari',
                                 '(33)HuquqMuhafizelcbari', '(34)DovletQulluqculcbari', '(35)DiplomatlarinIcbari',
                                 '(36)AuditorPeshe MesuliyyIcbari', '(37)IcbariDashinmazEmlak',
                                 '(38)Icbari DashinmazEmlakMesul', '(39)IcbariNVSMMS', '(40)IcbariSernishinFerdi Qeza',
                                 '(41)Sefer', 'titul sığortası', 'hüquqi xərclərin sığortası',
                                 'həyatın ölüm halından sığortası',
                                 'İstehsalatda bədbəxt hadisələr və peşə xəstəlikləri nəticəsində peşə əmək qabiliyyətinin itirilməsi hallarından icbari sığorta'],
               'Forma 8': [25, 80, 20, 50, 20, 35, 20, 20, 50, 90, 30, 20, 40, 20, 20, 20, 20, 20, 20, 20, 20, 40, 40,
                           50,
                           20, 20, 20, 50, 20, 20, 20, 20, 30, 30, 35, 20, 35, 20, 20, 25, 25],
               'Forma 12': [70, 40, 20, 30, 20, 40, 20, 25, 20, 20, 80, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20, 20,
                            20,
                            20, 20, 20, 40, 20, 20, 20, 20, 30, 20, 30, 20, 30, 20, 20, 25, 25]
               }

        avg_emsal = pd.DataFrame(avg)
        avg_emsal['Forma 8'] = avg_emsal['Forma 8'] / 100
        emsal = avg_emsal[avg_emsal['SıgortaSinifi'] == 'titul sığortası']['Forma 8']
        return emsal

    ###-----------

    def form8(df, dfs):
        quarterly_dates = generate_quarterly_range(input_date,year)

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
        pivot_df = shift_x(pivot_df)

        pivot_df.columns = range(1, len(pivot_df.columns) + 1)

        dfs['zerer_kvartal'] = dfs['Sığоrtа hаdisəsinin bаş vеrdiyi tаriх'].apply(last_day_of_quarter)
        dfs['odenis_kvartal'] = dfs['Subroqasiya gəlirinin daxil olduğu tarix'].apply(last_day_of_quarter)

        filtered_dfs = dfs[
            (dfs['zerer_kvartal'].isin(quarterly_dates)) &
            (dfs['odenis_kvartal'].isin(quarterly_dates))
            ]

        pivot_dfs = filtered_dfs.pivot_table(
            index='zerer_kvartal',
            columns='odenis_kvartal',
            values='Ödənilmiş subroqasiya gəlirinin məbləği',
            aggfunc='sum',
            fill_value=0
        ).reindex(index=quarterly_dates, columns=quarterly_dates, fill_value=0).cumsum(axis=1)

        pivot_dfs = shift_rows_left(pivot_dfs)
        pivot_dfs = shift_x(pivot_dfs)

        pivot_dfs.columns = range(1, len(pivot_dfs.columns) + 1)

        pivot_df = pivot_df.apply(pd.to_numeric, errors='coerce') - pivot_dfs.apply(pd.to_numeric, errors='coerce')
        pivot_df = shift_x(pivot_df)

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
                coef_row.append(round(numerator / denominator, 4))
            else:
                coef_row.append('X')
        coef_row.append('X')

        for i in range(len(coef_row)):
            check_up = 0
            if coef_row[i] == 0:
                check_up += 1

        if check_up > 0:
            pivot_df.loc[len(pivot_df) + 1] = evez_list[0]
        else:
            pivot_df.loc[len(pivot_df) + 1] = coef_row
        # new_row = [coef_row[i] + evez_list[i] if coef_row[i] == 0 else coef_row[i] for i in range(len(coef_row))]
        # #16
        ink_amil = pivot_df.iloc[-1][:-1].tolist()
        result_list = []

        for i in range(len(ink_amil)):
            product = 1
            for j in range(i, len(ink_amil)):
                try:
                    value = 1 if ink_amil[j] == 'X' else float(ink_amil[j])
                except ValueError:
                    print(f"Invalid value at index {j}: {ink_amil[j]} — skipping this value.")
                    value = 1

                product *= value

            result_list.append(round(product, 4))

        result_list.append(1)
        pivot_df.loc[len(pivot_df) + 1] = result_list

        # 17
        last_row = pivot_df.iloc[-1]
        gec_amil = []

        for value in last_row:
            if value != 0 and value != 'X':
                gec_amil.append(round((1 / value), 4))
            else:
                gec_amil.append('X')

        pivot_df.loc[len(pivot_df) + 1] = gec_amil
        # right_columns
        # tekrarsigorta
        new_column = pd.Series(form_7.iloc[:, 5].values[1:], index=pivot_df.index[:len(form_7) - 1]).fillna(0).infer_objects(copy=False)

        #pd.Series(form_7.iloc[:, 5].values[1:], index=pivot_df.index[:len(form_7) - 1]).fillna(0)
        pivot_df['Qazanılmış məcmu sığorta (təkrarsığorta) haqları'] = new_column.reindex(pivot_df.index).fillna('X')
        # zererler emsali
        fixed_value = pivot_df.iloc[-2, 11]  #
        result_list = []
        for i in range(len(form_7) - 1):
            last_value = pivot_df.iloc[i, -(i + 2)]
            col_13_value = pivot_df.iloc[i, -1]
            if col_13_value > 0:
                result = round(((last_value * fixed_value) / col_13_value), 4)
            else:
                result = 0
            result_list.append(result)
        result_list = pd.Series(result_list, index=pivot_df.index[:len(form_7) - 1])
        pivot_df['Ödənilmiş zərərlər əmsalı'] = result_list.reindex(pivot_df.index).fillna('X')
        # ortaqiymet
        mean_value = pivot_df.iloc[:, -1][:len(form_7) - 1].mean()
        if mean_value > 0:
            mean_list = [round(mean_value, 4)] * (len(form_7) - 1)
        else:
            mean_list = [float(avg_emsal.iloc[0])] * (len(form_7) - 1)
        mean_list = pd.Series(mean_list, index=pivot_df.index[:len(form_7) - 1])
        pivot_df['Ödənilmiş zərər əmsalının orta qiyməti'] = mean_list.reindex(pivot_df.index).fillna('X')
        # gozlenilen mebleg
        expect_list = []
        for val_13, val_15 in zip(pivot_df.iloc[:, -3][:len(form_7) - 1], pivot_df.iloc[:, -1][:len(form_7) - 1]):
            result = round(val_13 * val_15, 2)
            expect_list.append(result)
        expect_list = [round(x, 2) for x in expect_list]
        expect_list = pd.Series(expect_list, index=pivot_df.index[:len(form_7) - 1])
        pivot_df['Baş vermiş zərərlərin gözlənilən məbləği'] = expect_list.reindex(pivot_df.index).fillna('X')
        # odenilmemis mecmu mebleg
        fixx = pivot_df.iloc[-1, len(form_7) - 2]
        try:
            fixx = float(fixx)
        except ValueError:
            fixx = 0
        npaid = []
        for i in range(len(form_7) - 1):
            value = pivot_df.iloc[i, -1]
            value = float(value)
            result = round(((1 - fixx) * value), 2)
            npaid.append(result)
        npaid = pd.Series(npaid, index=pivot_df.index[:len(form_7) - 1])
        pivot_df['Baş vermiş, lakin ödənilməmiş zərərlərin məcmu məbləği'] = npaid.reindex(pivot_df.index).fillna('X')

        """
        fixx = pivot_df.iloc[-1, len(form_7) - 2]
        npaid = []
        for i in range(len(form_7) - 1):
            value = pivot_df.iloc[i, -1]
            result = round(((1 - fixx) * value), 2)
            npaid.append(result)
        npaid = pd.Series(npaid, index=pivot_df.index[:len(form_7) - 1])
        pivot_df['Baş vermiş, lakin ödənilməmiş zərərlərin məcmu məbləği'] = npaid.reindex(pivot_df.index).fillna('X')"""
        # tenzimlenmemis zerer
        tenzim_list = list(form_3.iloc[2:-1, -1])
        tenzim_list = pd.Series(tenzim_list, index=pivot_df.index[:len(form_3.index[1:-2])])
        pivot_df['Bildirilmiş, lakin tənzimlənməmiş zərərlərin məbləği'] = tenzim_list.reindex(pivot_df.index).fillna(
            'X')
        # bildirilmemis mecmu mebleg
        diff_list = (pivot_df.iloc[:, -2][:len(form_7) - 1].map(is_float) - pivot_df.iloc[:, -1][:len(form_7) - 1].map(
            is_float)).tolist()
        diff_list = [max(x, 0) for x in diff_list]
        diff_list = pd.Series(diff_list, index=pivot_df.index[:len(form_7) - 1])
        pivot_df['Baş vermiş, lakin bildirilməmiş zərərlərin məcmu məbləği'] = diff_list.reindex(pivot_df.index).fillna(
            'X')
        # rows
        x_list = ['X'] * (len(pivot_df.columns) - 1)
        sum_value = pivot_df.iloc[:len(form_7) - 1, -1].sum()
        sum_list = x_list + [round(sum_value, 2)]
        sum_list = pd.Series(sum_list, index=pivot_df.columns[:len(sum_list)])
        pivot_df.loc[len(pivot_df) + 1] = sum_list.fillna('X')

        bvbze = x_list + [round((sum_value * 1.03), 2)]
        bvbze = pd.Series(bvbze, index=pivot_df.columns[:len(bvbze)])
        pivot_df.loc[len(pivot_df) + 1] = bvbze.fillna('X')
        pivot_df = pivot_df.reset_index(drop=False)

        return pivot_df


    def form11(data):
        quarterly_dates = generate_quarterly_range(input_date,year)
        filtered_df = data[
            data['SigortaHadisesininBasverdiyiTarix'].apply(last_day_of_quarter).isin(quarterly_dates) &
            data['VerilmisOdenisTarixi'].isna() & 
            data['VerilmisSigortaOdenisi'].isna()
        ]
        
        filtered_df1 = data[
            (data['SigortaHadisesininBasverdiyiTarix'] <= quarterly_dates[0]) &
            data['VerilmisOdenisTarixi'].isna() & 
            data['VerilmisSigortaOdenisi'].isna()
        ]
        sorted_df = filtered_df.sort_values('SigortaHadisesininBasverdiyiTarix').reset_index(drop=True)
        new_df = pd.DataFrame(index=[''] + list(quarterly_dates))
        new_df['Sətrin kodu'] = [f'{i:02d}' for i in range(1, len(new_df) + 1)]
        new_df['Bildirilmiş, lakin tənzimlənməmiş zərərlərin hesabat tarixinə məbləği'] = (
            sorted_df.groupby(sorted_df['SigortaHadisesininBasverdiyiTarix'].apply(last_day_of_quarter))[
                'TekrarsigortacininBorcPayi'
            ].sum().reindex(new_df.index, fill_value=0)
        )
        new_df.iloc[0, 1] += filtered_df1['TekrarsigortacininBorcPayi'].sum()
        new_df['Bildirilmiş, lakin tənzimlənməmiş zərərlərdə təkrarsığortalının payı'] = 0
        new_df['Təkrarsığorta müqavilələrinə vaxtından əvvəl xitam verilməsi məbləği'] = (
            new_df.iloc[:, 1] + new_df.iloc[:, 2]) * 0.03
        new_df['Bildirilmiş zərərlər ehtiyatında təkrarsığortaçıların payı'] = new_df.iloc[:, 1:4].sum(axis=1)
        new_df.loc['Yekun BTZE'] = ['X'] + list(new_df.iloc[:, 1:].sum())
        new_df.index = [''] + [f'rüb {i}' for i in range(len(quarterly_dates), 0, -1)] + ['Yekun BTZE']

        return new_df.reset_index().rename(columns={'index': 'Sığorta hadisələrinin baş verdiyi rüblər'}) 

    def form12(df, dfs):
        quarterly_dates = generate_quarterly_range(input_date,year)

        df['zerer_kvartal'] = df['SigortaHadisesininBasverdiyiTarix'].apply(last_day_of_quarter)
        df['odenis_kvartal'] = df['VerilmisOdenisTarixi'].apply(last_day_of_quarter)

        filtered_df = df[
            (df['zerer_kvartal'].isin(quarterly_dates)) &
            (df['odenis_kvartal'].isin(quarterly_dates))
            ]

        pivot_df = filtered_df.pivot_table(
            index='zerer_kvartal',
            columns='odenis_kvartal',
            values='TekrarsigortacininPayi',
            aggfunc='sum',
            fill_value=0
        ).reindex(index=quarterly_dates, columns=quarterly_dates, fill_value=0).cumsum(axis=1)

        pivot_df = shift_rows_left(pivot_df)
        pivot_df = shift_x(pivot_df)

        pivot_df.columns = range(1, len(pivot_df.columns) + 1)

        dfs['zerer_kvartal'] = dfs['Sığоrtа hаdisəsinin bаş vеrdiyi tаriх'].apply(last_day_of_quarter)
        dfs['odenis_kvartal'] = dfs['Subroqasiya gəlirinin daxil olduğu tarix'].apply(last_day_of_quarter)

        filtered_dfs = dfs[
            (dfs['zerer_kvartal'].isin(quarterly_dates)) &
            (dfs['odenis_kvartal'].isin(quarterly_dates))
            ]

        pivot_dfs = filtered_dfs.pivot_table(
            index='zerer_kvartal',
            columns='odenis_kvartal',
            values='Ödənilmiş subroqasiya gəlirinin məbləği',
            aggfunc='sum',
            fill_value=0
        ).reindex(index=quarterly_dates, columns=quarterly_dates, fill_value=0).cumsum(axis=1)

        pivot_dfs = shift_rows_left(pivot_dfs)
        pivot_dfs = shift_x(pivot_dfs)

        pivot_dfs.columns = range(1, len(pivot_dfs.columns) + 1)

        pivot_df = pivot_df.apply(pd.to_numeric, errors='coerce') - pivot_dfs.apply(pd.to_numeric, errors='coerce')
        pivot_df = shift_x(pivot_df)

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
                coef_row.append(round(numerator / denominator, 4))
            else:
                coef_row.append('X')
        coef_row.append('X')
        for i in range(len(coef_row)):
            check_up = 0
            if coef_row[i] == 0:
                check_up += 1

        if check_up > 0:
            pivot_df.loc[len(pivot_df) + 1] = evez_list[1]
        else:
            pivot_df.loc[len(pivot_df) + 1] = coef_row

        # 16
        ink_amil = pivot_df.iloc[-1][:-1].tolist()
        result_list = []

        for i in range(len(ink_amil)):
            product = 1
            for j in range(i, len(ink_amil)):
                try:
                    value = 1 if ink_amil[j] == 'X' else float(ink_amil[j])
                except ValueError:
                    print(f"Invalid value at index {j}: {ink_amil[j]} — skipping this value.")
                    value = 1

                product *= value

            result_list.append(round(product, 4))

        result_list.append(1)
        pivot_df.loc[len(pivot_df) + 1] = result_list

        # 17
        last_row = pivot_df.iloc[-1]
        gec_amil = []

        for value in last_row:
            if value != 0 and value != 'X':
                gec_amil.append(round((1 / value), 4))
            else:
                gec_amil.append('X')

        pivot_df.loc[len(pivot_df) + 1] = gec_amil
        # right_columns
        # tekrarsigorta
        new_column = pd.Series(form_7.iloc[:, 5].values[1:], index=pivot_df.index[:len(form_7) - 1]).fillna(0)
        pivot_df['Qazanılmış məcmu sığorta (təkrarsığorta) haqları'] = new_column.reindex(pivot_df.index).fillna('X')
        # zererler emsali
        fixed_value = pivot_df.iloc[-2, 11]  #
        result_list = []
        for i in range(len(form_7) - 1):
            last_value = pivot_df.iloc[i, -(i + 2)]
            col_13_value = pivot_df.iloc[i, -1]
            if col_13_value > 0:
                result = round(((last_value * fixed_value) / col_13_value), 4)
            else:
                result = 0
            result_list.append(result)
        result_list = pd.Series(result_list, index=pivot_df.index[:len(form_7) - 1])
        pivot_df['Ödənilmiş zərərlər əmsalı'] = result_list.reindex(pivot_df.index).fillna('X')
        # ortaqiymet
        mean_value = pivot_df.iloc[:, -1][:len(form_7) - 1].mean()
        if mean_value > 0:
            mean_list = [round(mean_value, 4)] * (len(form_7) - 1)
        else:
            mean_list = [float(avg_emsal)] * (len(form_7) - 1)
        mean_list = pd.Series(mean_list, index=pivot_df.index[:len(form_7) - 1])
        pivot_df['Ödənilmiş zərər əmsalının orta qiyməti'] = mean_list.reindex(pivot_df.index).fillna('X')
        # gozlenilen mebleg
        expect_list = []
        for val_13, val_15 in zip(pivot_df.iloc[:, -3][:len(form_7) - 1], pivot_df.iloc[:, -1][:len(form_7) - 1]):
            result = round(val_13 * val_15, 2)
            expect_list.append(result)
        expect_list = [round(x, 2) for x in expect_list]
        expect_list = pd.Series(expect_list, index=pivot_df.index[:len(form_7) - 1])
        pivot_df['Baş vermiş zərərlərin gözlənilən məbləği'] = expect_list.reindex(pivot_df.index).fillna('X')
        # odenilmemis mecmu mebleg
        fixx = pivot_df.iloc[-1, len(form_7) - 2]
        try:
            fixx = float(fixx)
        except ValueError:
            fixx = 0
        npaid = []
        for i in range(len(form_7) - 1):
            value = pivot_df.iloc[i, -1]
            value = float(value)
            result = round(((1 - fixx) * value), 2)
            npaid.append(result)
        npaid = pd.Series(npaid, index=pivot_df.index[:len(form_7) - 1])
        pivot_df['Baş vermiş, lakin ödənilməmiş zərərlərin məcmu məbləği'] = npaid.reindex(pivot_df.index).fillna('X')
        # tenzimlenmemis zerer
        tenzim_list = list(form_3.iloc[2:-1, -1])
        tenzim_list = pd.Series(tenzim_list, index=pivot_df.index[:len(form_3.index[1:-2])])
        pivot_df['Bildirilmiş, lakin tənzimlənməmiş zərərlərin məbləği'] = tenzim_list.reindex(pivot_df.index).fillna(
            'X')
        # bildirilmemis mecmu mebleg
        diff_list = (pivot_df.iloc[:, -2][:len(form_7) - 1].map(is_float) - pivot_df.iloc[:, -1][:len(form_7) - 1].map(
            is_float)).tolist()
        diff_list = [max(x, 0) for x in diff_list]
        diff_list = pd.Series(diff_list, index=pivot_df.index[:len(form_7) - 1])
        pivot_df['Baş vermiş, lakin bildirilməmiş zərərlərin məcmu məbləği'] = diff_list.reindex(pivot_df.index).fillna(
            'X')

        # rows
        x_list = ['X'] * (len(pivot_df.columns) - 1)
        sum_value = pivot_df.iloc[:len(form_7) - 1, -1].sum()
        sum_list = x_list + [round(sum_value, 2)]
        sum_list = pd.Series(sum_list, index=pivot_df.columns[:len(sum_list)])
        pivot_df.loc[len(pivot_df) + 1] = sum_list.fillna('X')

        bvbze = x_list + [round((sum_value * 1.03), 2)]
        bvbze = pd.Series(bvbze, index=pivot_df.columns[:len(bvbze)])
        pivot_df.loc[len(pivot_df) + 1] = bvbze.fillna('X')
        pivot_df = pivot_df.reset_index(drop=False)
        return pivot_df

    def form9():
        value_1 = form_8.iloc[-1, -1]
        value_2 = (form_3.iloc[-1, -1] * 0.25).round(2)
        value_3 = (0.025 * form_7.iloc[-4:, 6].sum()).round(2)

        max_value = max(value_1, value_2, value_3)

        data = {
            'Sıra No': [1, 2, 3, 4],
            'Göstəriciləri adı': [
                'Hesabat tarixinə BVBZE-nin üçbucaq metodu ilə hesablanmış məbləği',
                'Hesabat tarixinə BTZE-nin 25 %-i',
                'Hesabat dövrü ərzində qazanılmış məcburi sığorta (təkrarsığorta) haqlarının 2,5%-i',
                'Hesabat tarixinə BVBZE ("1000", "1100" və "1200"-dən böyük olanı)'],
            'Sətirin kodu': [1000, 1100, 1200, 1300],
            'Məbləğ': [
                value_1,
                value_2,
                value_3,
                max_value]}

        result_df = pd.DataFrame(data)
        return result_df

    def form13():
        value_1 = form_12.iloc[-1, -1]
        value_2 = (form_11.iloc[-1, -1] * 0.25).round(2)
        value_3 = (0.025 * form_10.iloc[-4:, 6].sum()).round(2)

        max_value = max(value_1, value_2, value_3)

        data = {
            'Sıra No': [1, 2, 3, 4],
            'Göstəriciləri adı': [
                'Hesabat tarixinə BVBZE-də təkrarsığortaçıların payının üçbucaq metodu ilə hesablanmış məbləği',
                'Hesabat tarixinə BTZE-də təkrarsığortaçıların payının 25 %-i',
                'Hesabat dövrü ərzində qazanılmış məcmu təkrarsığorta haqlarının 2,5%-i',
                'Hesabat tarixinə BVBZE-də təkrarsığortaçıların payı ("2000", "2100" və "2200"-dən böyük olanı)'],
            'Sətirin kodu': [2000, 2100, 2200, 2300],
            'Məbləğ': [
                value_1,
                value_2,
                value_3,
                max_value]}

        result_df = pd.DataFrame(data)
        return result_df

    # Xitam verilmiş müqavilələri gətir
    jurnals = active_date(df, dz, dfs, sinif, input_date)

    evez_list = emsal(sinif)
    avg_emsal = avgemsal()
    formatted_date = tarix_formatla(input_date)
    filtered_a = prepare_insurance(jurnals[0], 'ALINAN-FAKULTATİV TƏKRAR SIĞORTA', invert=True)
    filtered_b = prepare_insurance(jurnals[0], 'ALINAN-FAKULTATİV TƏKRAR SIĞORTA')
    form10_4a = end_quarter10(jurnals[0])
    form1_a = form1(filtered_a, columns_form1_a)
    form1_b = form1(filtered_b, columns_form1_b)
    form2_a = form2(filtered_a, columns_form2_a, input_date)
    form2_b = form2(filtered_b, columns_form2_b, input_date)
    form_2yekun = form2yekun(form2_a, form2_b)
    form_4 = [form4(prepare_class(jurnals[0], grp)) for grp in insurance_groups]
    form_4yekun = form4yekun()
    form_5 = [form5(prepare_class(jurnals[0], grp)) for grp in insurance_groups]
    form_5yekun = form5yekun()
    form_6 = form6()
    form7_1a = end_quarter(jurnals[0])
    form_7 = form7(db)
    form_10 = form10(dq)
    form_3 = form3(jurnals[1])
    form_8 = form8(jurnals[1], jurnals[2])
    form_9 = form9()
    form_11 = form11(jurnals[1])
    form_12 = form12(jurnals[1], jurnals[2])
    form_13 = form13()

    return [formatted_date, form1_a, form1_b, form2_a, form2_b,
            form_2yekun, form_4, form_4yekun,
            form_5, form_5yekun, form_6, form_7, form_10,
            form_3, form_8, form_9, form_11, form_12, form_13]


def group_insurance(data):
    unique_values = data['SigortaSinifi'].dropna().unique()
    result_df = pd.DataFrame(unique_values, columns=['SigortaSinifi'])
    class_data = {
        'SigortaSinifi': [
            "(01)FerdiQeza", "(02)Tibbi", "(03)EmlakYanginDigerRisk", "(04)AvtoKasko", "(05)DemiryolNeqliyyVasitesi",
            "(06)HavaNeqliyyKasko", "(07)SuNeqliyyKasko", "(08)Yuk", "(09)KendTeserrufBitki", "(10)KendTeserrufHeyvan",
            "(11)IshcilerinDeleduzlug", "(12)PulvePulSenedSaxtalash", "(13)AvtoKonulluMesuliyy",
            "(14)DemiryolNeqliySahibMesuliyy",
            "(15)HavaNeqliySahibMesuliyy", "(16)SuNeqliySahibMesuliyy", "(17)YukDashiyanMesuliyy",
            "(18)MulkiMuqavileUzreMesuliyy",
            "(19)PesheMesuliyy", "(20)IshegoturenMesuliyy", "(21)UmumiMulkiMesuliyy", "(22)Kredit", "(23)Ipoteka",
            "(24)EmlakinDeyerdenDushmesi", "(25)IshinDayanmasiRiski", "(26)AvtoIcbariMesuliyy", "(27)SernishinIcbari",
            "(28)IcbariEkoloji", "(29)YanginIcbari", "(30)DeputatlarinIcbari", "(31)TibbiPersonalinAIDSden",
            "(32)HerbiQulluqcularinIcbari",
            "(33)HuquqMuhafizeIcbari", "(34)DovletQulluqcuIcbari", "(35)DiplomatlarinIcbari",
            "(36)AuditorPesheMesuliyyIcbari",
            "(37)IcbariDashinmazEmlak", "(38)IcbariDashinmazEmlakMesul", "(39)IcbariNVSMMS",
            "(40)IcbariSernishinFerdiQeza", "(41)Sefer"
        ],
        'Kategoriya': [
            "B", "B", "B", "B", "B",
            "B", "B", "B", "B", "B",
            "B", "B", "A", "A",
            "A", "A", "A", "A",
            "B", "B", "A", "B", "B",
            "B", "B", "A", "B",
            "B", "B", "B", "B", "B",
            "B", "B", "B", "B",
            "A", "A", "A", "B", "B"
        ]
    }
    dff = pd.DataFrame(class_data)
    result_df = result_df.merge(dff, left_on='SigortaSinifi', right_on='SigortaSinifi')
    a_group = list(result_df['SigortaSinifi'][result_df['Kategoriya'] == 'A'])
    b_group = list(result_df['SigortaSinifi'][result_df['Kategoriya'] == 'B'])

    return [a_group, b_group]
