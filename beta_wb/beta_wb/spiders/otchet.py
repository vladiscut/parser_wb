import pandas as pd
import re
import time
from styleframe import StyleFrame, Styler, utils
import xlsxwriter
import math


def Otchet():
    df_now_odezda = pd.read_excel(r'all_odezhda.xlsx',converters={'Артикул':int, 'Отзывов':int,'Купили раз':int })
    df_now_obuv = pd.read_excel(r'all_obuv.xlsx',converters={'Артикул':int, 'Отзывов':int,'Купили раз':int })
    df_now_sumki = pd.read_excel(r'all_sumki.xlsx',converters={'Артикул':int, 'Отзывов':int,'Купили раз':int })
    df_now_igrushki = pd.read_excel(r'all_igrushki.xlsx',converters={'Артикул':int, 'Отзывов':int,'Купили раз':int })
    df_now_uvelirka = pd.read_excel(r'all_uvelirka.xlsx',converters={'Артикул':int, 'Отзывов':int,'Купили раз':int })

    df_last = pd.read_excel(r'all.xlsx',converters={'Артикул':int, 'Отзывов':int,'Купили раз':int })
    df_buyer = pd.read_excel('байер.xlsb',dtype={'Код артикула WB':int },engine='pyxlsb',skiprows = 10)

    df_now_odezda= df_now_odezda.merge(df_buyer[['Артикул_kari','Баер','Код артикула WB']], left_on='Артикул',right_on = 'Код артикула WB', how='left')
    df_now_obuv= df_now_obuv.merge(df_buyer[['Артикул_kari','Баер','Код артикула WB']], left_on='Артикул',right_on = 'Код артикула WB', how='left')
    df_now_sumki= df_now_sumki.merge(df_buyer[['Артикул_kari','Баер','Код артикула WB']], left_on='Артикул',right_on = 'Код артикула WB', how='left')
    df_now_igrushki= df_now_igrushki.merge(df_buyer[['Артикул_kari','Баер','Код артикула WB']], left_on='Артикул',right_on = 'Код артикула WB', how='left')
    df_now_uvelirka= df_now_uvelirka.merge(df_buyer[['Артикул_kari','Баер','Код артикула WB']], left_on='Артикул',right_on = 'Код артикула WB', how='left')

    df_now_odezda['Артикул Кари'] = df_now_odezda['Артикул_kari']
    df_now_odezda['Байер'] = df_now_odezda['Баер']
    df_now_odezda = df_now_odezda.drop(columns=['Артикул_kari', 'Баер', 'Код артикула WB'])

    df_now_obuv['Артикул Кари'] = df_now_obuv['Артикул_kari']
    df_now_obuv['Байер'] = df_now_obuv['Баер']
    df_now_obuv = df_now_obuv.drop(columns=['Артикул_kari', 'Баер', 'Код артикула WB'])

    df_now_sumki['Артикул Кари'] = df_now_sumki['Артикул_kari']
    df_now_sumki['Байер'] = df_now_sumki['Баер']
    df_now_sumki = df_now_sumki.drop(columns=['Артикул_kari', 'Баер', 'Код артикула WB'])

    df_now_igrushki['Артикул Кари'] = df_now_igrushki['Артикул_kari']
    df_now_igrushki['Байер'] = df_now_igrushki['Баер']
    df_now_igrushki = df_now_igrushki.drop(columns=['Артикул_kari', 'Баер', 'Код артикула WB'])

    df_now_uvelirka['Артикул Кари'] = df_now_uvelirka['Артикул_kari']
    df_now_uvelirka['Байер'] = df_now_uvelirka['Баер']
    df_now_uvelirka = df_now_uvelirka.drop(columns=['Артикул_kari', 'Баер', 'Код артикула WB'])


    df_now_odezda= df_now_odezda.merge(df_last[['Артикул','Отзывов','Купили раз']], on='Артикул', how='left')
    df_now_obuv= df_now_obuv.merge(df_last[['Артикул','Отзывов','Купили раз']], on='Артикул', how='left')
    df_now_sumki= df_now_sumki.merge(df_last[['Артикул','Отзывов','Купили раз']], on='Артикул', how='left')
    df_now_igrushki= df_now_igrushki.merge(df_last[['Артикул','Отзывов','Купили раз']], on='Артикул', how='left')
    df_now_uvelirka= df_now_uvelirka.merge(df_last[['Артикул','Отзывов','Купили раз']], on='Артикул', how='left')

    df_now_odezda['+ к отзывам'] = df_now_odezda['Отзывов_x'] - df_now_odezda['Отзывов_y']
    df_now_odezda['+к купили раз'] = df_now_odezda['Купили раз_x'] - df_now_odezda['Купили раз_y']

    df_now_obuv['+ к отзывам'] = df_now_obuv['Отзывов_x'] - df_now_obuv['Отзывов_y']
    df_now_obuv['+к купили раз'] = df_now_obuv['Купили раз_x'] - df_now_obuv['Купили раз_y']

    df_now_sumki['+ к отзывам'] = df_now_sumki['Отзывов_x'] - df_now_sumki['Отзывов_y']
    df_now_sumki['+к купили раз'] = df_now_sumki['Купили раз_x'] - df_now_sumki['Купили раз_y']

    df_now_igrushki['+ к отзывам'] = df_now_igrushki['Отзывов_x'] - df_now_igrushki['Отзывов_y']
    df_now_igrushki['+к купили раз'] = df_now_igrushki['Купили раз_x'] - df_now_igrushki['Купили раз_y']

    df_now_uvelirka['+ к отзывам'] = df_now_uvelirka['Отзывов_x'] - df_now_uvelirka['Отзывов_y']
    df_now_uvelirka['+к купили раз'] = df_now_uvelirka['Купили раз_x'] - df_now_uvelirka['Купили раз_y']

    for i in range(len(df_now_odezda)):
        if (df_now_odezda['+ к отзывам'][i] <= 0):
            df_now_odezda['+ к отзывам'][i] = float('nan')
        if (df_now_odezda['+к купили раз'][i] <= 0):
            df_now_odezda['+к купили раз'][i] = float('nan')

    for i in range(len(df_now_obuv)):
        if (df_now_obuv['+ к отзывам'][i] <= 0):
            df_now_obuv['+ к отзывам'][i] = float('nan')
        if (df_now_obuv['+к купили раз'][i] <= 0):
            df_now_obuv['+к купили раз'][i] = float('nan')

    for i in range(len(df_now_sumki)):
        if (df_now_sumki['+ к отзывам'][i] <= 0):
            df_now_sumki['+ к отзывам'][i] = float('nan')
        if (df_now_sumki['+к купили раз'][i] <= 0):
            df_now_sumki['+к купили раз'][i] = float('nan')

    for i in range(len(df_now_igrushki)):
        if (df_now_igrushki['+ к отзывам'][i] <= 0):
            df_now_igrushki['+ к отзывам'][i] = float('nan')
        if (df_now_igrushki['+к купили раз'][i] <= 0):
            df_now_igrushki['+к купили раз'][i] = float('nan')

    for i in range(len(df_now_uvelirka)):
        if (df_now_uvelirka['+ к отзывам'][i] <= 0):
            df_now_uvelirka['+ к отзывам'][i] = float('nan')
        if (df_now_uvelirka['+к купили раз'][i] <= 0):
            df_now_uvelirka['+к купили раз'][i] = float('nan')


    df_now_odezda = df_now_odezda.rename(columns={"Отзывов_x": "Отзывов", "Купили раз_x": "Купили раз"})
    df_now_odezda = df_now_odezda.drop(columns=['Отзывов_y', 'Купили раз_y'])

    df_now_obuv = df_now_obuv.rename(columns={"Отзывов_x": "Отзывов", "Купили раз_x": "Купили раз"})
    df_now_obuv = df_now_obuv.drop(columns=['Отзывов_y', 'Купили раз_y'])

    df_now_sumki = df_now_sumki.rename(columns={"Отзывов_x": "Отзывов", "Купили раз_x": "Купили раз"})
    df_now_sumki = df_now_sumki.drop(columns=['Отзывов_y', 'Купили раз_y'])

    df_now_igrushki = df_now_igrushki.rename(columns={"Отзывов_x": "Отзывов", "Купили раз_x": "Купили раз"})
    df_now_igrushki = df_now_igrushki.drop(columns=['Отзывов_y', 'Купили раз_y'])

    df_now_uvelirka = df_now_uvelirka.rename(columns={"Отзывов_x": "Отзывов", "Купили раз_x": "Купили раз"})
    df_now_uvelirka = df_now_uvelirka.drop(columns=['Отзывов_y', 'Купили раз_y'])

    df_now_odezda = df_now_odezda.drop_duplicates()
    df_now_obuv = df_now_obuv.drop_duplicates()
    df_now_sumki = df_now_sumki.drop_duplicates()
    df_now_igrushki = df_now_igrushki.drop_duplicates()
    df_now_uvelirka = df_now_uvelirka.drop_duplicates()

    df_pivot_obuv = df_now_obuv.pivot_table(index=["Направление", "Группа", "Подгруппа"], values=["Артикул","Артикул Кари"], aggfunc='count')
    df_pivot_odezda = df_now_odezda.pivot_table(index=["Направление", "Группа", "Подгруппа"], values=["Артикул","Артикул Кари"], aggfunc='count')
    df_pivot_sumki = df_now_sumki.pivot_table(index=["Направление", "Группа", "Подгруппа"], values=["Артикул","Артикул Кари"], aggfunc='count')
    df_pivot_igrushki = df_now_igrushki.pivot_table(index=["Направление", "Группа", "Подгруппа"], values=["Артикул","Артикул Кари"], aggfunc='count')
    df_pivot_uvelirka = df_now_uvelirka.pivot_table(index=["Направление", "Группа", "Подгруппа"], values=["Артикул","Артикул Кари"], aggfunc='count')

    excel_dir = r".\ready\парсинг_2021_18_одежда.xlsx"

    with pd.ExcelWriter(excel_dir, engine='xlsxwriter') as writer:
        df_now_odezda.to_excel(writer, 'Парсинг', index=False)
        df_pivot_odezda.to_excel(writer, 'Сводная таблица')
        workbook = writer.book
        worksheet = writer.sheets['Сводная таблица']

        format1 = workbook.add_format({'text_wrap': True, 'valign': 'vcenter', 'fg_color': '#D7E3BC', 'border': 1})
        format2 = workbook.add_format({'text_wrap': True, 'valign': 'vcenter', 'fg_color': '#D7E4BC', 'border': 1})

        worksheet.set_column('A:A', 25, format1)
        worksheet.set_column('B:B', 25, format2)
        worksheet.set_column('C:C', 20, format2)
        worksheet.set_column('D:D', 13, format2)
        worksheet.set_column('E:E', 13, format2)

        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'fg_color': '#FFD700',
            'border': 1})

        cols = ['Направление', 'Группа', 'Подгруппа', 'Артикул', 'Артикул Кари']

        # Write the column headers with the defined format.
        for col_num, value in enumerate(cols):
            worksheet.write(0, col_num, value, header_format)

        worksheet = writer.sheets['Парсинг']

        format1 = workbook.add_format(
            {'text_wrap': False, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#D7E3BC', 'border': 1})
        format2 = workbook.add_format(
            {'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#D7E4BC', 'border': 1})

        worksheet.set_column('A:Q', 10, format1)
        #     worksheet.set_column('B:B', 25, format2)
        worksheet.set_column('C:C', 18, format1)
        worksheet.set_column('D:D', 22, format1)
        #     worksheet.set_column('E:E', 13, format2)

        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'fg_color': '#FFD700',
            'border': 1})

        cols = df_now_odezda.columns.tolist()

        # Write the column headers with the defined format.
        for col_num, value in enumerate(cols):
            worksheet.write(0, col_num, value, header_format)
        writer.save()

    excel_dir = r".\ready\парсинг_2021_18_обувь.xlsx"

    with pd.ExcelWriter(excel_dir, engine='xlsxwriter') as writer:
        df_now_obuv.to_excel(writer, 'Парсинг', index=False)
        df_pivot_obuv.to_excel(writer, 'Сводная таблица')
        workbook = writer.book
        worksheet = writer.sheets['Сводная таблица']

        format1 = workbook.add_format({'text_wrap': True, 'valign': 'vcenter', 'fg_color': '#D7E3BC', 'border': 1})
        format2 = workbook.add_format({'text_wrap': True, 'valign': 'vcenter', 'fg_color': '#D7E4BC', 'border': 1})

        worksheet.set_column('A:A', 25, format1)
        worksheet.set_column('B:B', 25, format2)
        worksheet.set_column('C:C', 20, format2)
        worksheet.set_column('D:D', 13, format2)
        worksheet.set_column('E:E', 13, format2)

        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'fg_color': '#FFD700',
            'border': 1})

        cols = ['Направление', 'Группа', 'Подгруппа', 'Артикул', 'Артикул Кари']

        # Write the column headers with the defined format.
        for col_num, value in enumerate(cols):
            worksheet.write(0, col_num, value, header_format)

        worksheet = writer.sheets['Парсинг']

        format1 = workbook.add_format(
            {'text_wrap': False, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#D7E3BC', 'border': 1})
        format2 = workbook.add_format(
            {'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#D7E4BC', 'border': 1})

        worksheet.set_column('A:Q', 10, format1)
        #     worksheet.set_column('B:B', 25, format2)
        worksheet.set_column('C:C', 18, format1)
        worksheet.set_column('D:D', 22, format1)
        #     worksheet.set_column('E:E', 13, format2)

        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'fg_color': '#FFD700',
            'border': 1})

        cols = df_now_obuv.columns.tolist()

        # Write the column headers with the defined format.
        for col_num, value in enumerate(cols):
            worksheet.write(0, col_num, value, header_format)
        writer.save()

    excel_dir = r".\ready\парсинг_2021_17_сумки.xlsx"

    with pd.ExcelWriter(excel_dir, engine='xlsxwriter') as writer:
        df_now_sumki.to_excel(writer, 'Парсинг', index=False)
        df_pivot_sumki.to_excel(writer, 'Сводная таблица')
        workbook = writer.book
        worksheet = writer.sheets['Сводная таблица']

        format1 = workbook.add_format({'text_wrap': True, 'valign': 'vcenter', 'fg_color': '#D7E3BC', 'border': 1})
        format2 = workbook.add_format({'text_wrap': True, 'valign': 'vcenter', 'fg_color': '#D7E4BC', 'border': 1})

        worksheet.set_column('A:A', 25, format1)
        worksheet.set_column('B:B', 25, format2)
        worksheet.set_column('C:C', 20, format2)
        worksheet.set_column('D:D', 13, format2)
        worksheet.set_column('E:E', 13, format2)

        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'fg_color': '#FFD700',
            'border': 1})

        cols = ['Направление', 'Группа', 'Подгруппа', 'Артикул', 'Артикул Кари']

        # Write the column headers with the defined format.
        for col_num, value in enumerate(cols):
            worksheet.write(0, col_num, value, header_format)

        worksheet = writer.sheets['Парсинг']

        format1 = workbook.add_format(
            {'text_wrap': False, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#D7E3BC', 'border': 1})
        format2 = workbook.add_format(
            {'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#D7E4BC', 'border': 1})

        worksheet.set_column('A:Q', 10, format1)
        #     worksheet.set_column('B:B', 25, format2)
        worksheet.set_column('C:C', 18, format1)
        worksheet.set_column('D:D', 22, format1)
        #     worksheet.set_column('E:E', 13, format2)

        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'fg_color': '#FFD700',
            'border': 1})

        cols = df_now_sumki.columns.tolist()

        # Write the column headers with the defined format.
        for col_num, value in enumerate(cols):
            worksheet.write(0, col_num, value, header_format)
        writer.save()

    excel_dir = r".\ready\парсинг_2021_17_игрушки.xlsx"

    with pd.ExcelWriter(excel_dir, engine='xlsxwriter') as writer:
        df_now_igrushki.to_excel(writer, 'Парсинг', index=False)
        df_pivot_igrushki.to_excel(writer, 'Сводная таблица')
        workbook = writer.book
        worksheet = writer.sheets['Сводная таблица']

        format1 = workbook.add_format({'text_wrap': True, 'valign': 'vcenter', 'fg_color': '#D7E3BC', 'border': 1})
        format2 = workbook.add_format({'text_wrap': True, 'valign': 'vcenter', 'fg_color': '#D7E4BC', 'border': 1})

        worksheet.set_column('A:A', 25, format1)
        worksheet.set_column('B:B', 25, format2)
        worksheet.set_column('C:C', 20, format2)
        worksheet.set_column('D:D', 13, format2)
        worksheet.set_column('E:E', 13, format2)

        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'fg_color': '#FFD700',
            'border': 1})

        cols = ['Направление', 'Группа', 'Подгруппа', 'Артикул', 'Артикул Кари']

        # Write the column headers with the defined format.
        for col_num, value in enumerate(cols):
            worksheet.write(0, col_num, value, header_format)

        worksheet = writer.sheets['Парсинг']

        format1 = workbook.add_format(
            {'text_wrap': False, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#D7E3BC', 'border': 1})
        format2 = workbook.add_format(
            {'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#D7E4BC', 'border': 1})

        worksheet.set_column('A:Q', 10, format1)
        #     worksheet.set_column('B:B', 25, format2)
        worksheet.set_column('C:C', 18, format1)
        worksheet.set_column('D:D', 22, format1)
        #     worksheet.set_column('E:E', 13, format2)

        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'fg_color': '#FFD700',
            'border': 1})

        cols = df_now_igrushki.columns.tolist()

        # Write the column headers with the defined format.
        for col_num, value in enumerate(cols):
            worksheet.write(0, col_num, value, header_format)
        writer.save()

    excel_dir = r".\ready\парсинг2_2021_17_ювелирные.xlsx"

    with pd.ExcelWriter(excel_dir, engine='xlsxwriter') as writer:
        df_now_uvelirka.to_excel(writer, 'Парсинг', index=False)
        df_pivot_uvelirka.to_excel(writer, 'Сводная таблица')
        workbook = writer.book
        worksheet = writer.sheets['Сводная таблица']

        format1 = workbook.add_format({'text_wrap': True, 'valign': 'vcenter', 'fg_color': '#D7E3BC', 'border': 1})
        format2 = workbook.add_format({'text_wrap': True, 'valign': 'vcenter', 'fg_color': '#D7E4BC', 'border': 1})

        worksheet.set_column('A:A', 25, format1)
        worksheet.set_column('B:B', 25, format2)
        worksheet.set_column('C:C', 20, format2)
        worksheet.set_column('D:D', 13, format2)
        worksheet.set_column('E:E', 13, format2)

        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'fg_color': '#FFD700',
            'border': 1})

        cols = ['Направление', 'Группа', 'Подгруппа', 'Артикул', 'Артикул Кари']

        # Write the column headers with the defined format.
        for col_num, value in enumerate(cols):
            worksheet.write(0, col_num, value, header_format)

        worksheet = writer.sheets['Парсинг']

        format1 = workbook.add_format(
            {'text_wrap': False, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#D7E3BC', 'border': 1})
        format2 = workbook.add_format(
            {'text_wrap': True, 'align': 'center', 'valign': 'vcenter', 'fg_color': '#D7E4BC', 'border': 1})

        worksheet.set_column('A:Q', 10, format1)
        #     worksheet.set_column('B:B', 25, format2)
        worksheet.set_column('C:C', 18, format1)
        worksheet.set_column('D:D', 22, format1)
        #     worksheet.set_column('E:E', 13, format2)

        header_format = workbook.add_format({
            'bold': True,
            'text_wrap': True,
            'valign': 'vcenter',
            'align': 'center',
            'fg_color': '#FFD700',
            'border': 1})

        cols = df_now_uvelirka.columns.tolist()

        # Write the column headers with the defined format.
        for col_num, value in enumerate(cols):
            worksheet.write(0, col_num, value, header_format)
        writer.save()



