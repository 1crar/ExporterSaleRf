"""PDF and document parsers used before writing data to 1C."""

import glob
import os
import re

import camelot
import numpy as np
import pandas as pd
import pdfplumber
import PyPDF2
from camelot import read_pdf

from ..bitrix import send_msg_to_bot
from ..pdf.extract import get_info
from .utils import convDate, convDateTo1CFormat, convSum

# The legacy parser referenced dealID from module globals inside error paths.
# app.main updates this value before each parsed deal to preserve that behavior.
dealID = None

def extractBitrixDocInfo(myFile):
    with open(myFile, 'rb') as f:
        reader = PyPDF2.PdfReader(f)
        text = reader.pages[0].extract_text()
        #print(text)
        pagesCount = len(reader.pages)

    if "ВЭД ПАРТНЕР" in text:
        orgName = "ВЭД ПАРТНЕР ТОО"
    elif "ТРИАЛ" in text:
        orgName = "ТРИАЛ-ПВ ТОО"
    else:
        orgName = "ВЭД Партнер-Алматы ТОО"

    #print(myFile)
    type = 1
    resTables = camelot.read_pdf(filepath=myFile, line_scale=100, flavor='lattice')
    if not resTables or not ("ЗАКЛЮЧЕНИЕ ОБ ИДЕНТИФИКАЦИИ" in text):
        type = 2
        pagesStr = ",".join([str(i) for i in range(2, pagesCount + 1)])
        resTables = read_pdf(myFile, flavor="lattice", pages=pagesStr)

    items = []
    itemsCodes = []

    for table in resTables:
        #print(table)
        table = table.df.replace('', np.nan)
        table = table.dropna(how='all')
        #print(table)
        if type == 2:
            itemCol = table[1]
            itemCodeCol = table[3]
            try:
                items.extend(itemCol.iloc[2:].to_list())
                itemsCodes.extend(itemCodeCol.iloc[2:].to_list())
            except:
                continue
        else:
            itemCol = table[2]
            itemCodeCol = table[0]

            items.extend(itemCol.iloc[1:].to_list())
            itemsCodes.extend(itemCodeCol.iloc[1:].to_list())
            break
    #items = list(filter(lambda x: not np.isnan(x), items))
    #itemsCodes = list(filter(lambda x: not np.isnan(x), itemsCodes))
    #print('7834783734')
    #print(itemsCodes)
    i = 0
    '''while i < len(items) - 1:
        if type(items[i]) is not str or type(itemsCodes[i]) is not str or\
            len(itemsCodes[i]) < 10 or type(itemsCodes[i + 1]) is not str:

            if type(items[i + 1]) is str:
                items[i] = items[i + 1] if type(items[i]) is not str else (items[i] + '\n' + items[i + 1])
            if type(itemsCodes[i]) is str and len(itemsCodes[i]) != 10:
                itemsCodes[i] = itemsCodes[i + 1] if type(itemsCodes[i]) is not str else (itemsCodes[i] + itemsCodes[i + 1])

            items.remove(items[i + 1])
            itemsCodes.remove(itemsCodes[i + 1])
        else:
            i += 1'''
    print(itemsCodes)
    print(items)
    for i in range(len(items)):
        #print(items[i])
        try:
            if not items[i] or pd.isna(items[i]):
                items.remove(items[i])
            else:
                items[i] = items[i].replace('\n', ' ')
        except:
            break
    for i in range(len(itemsCodes)):
        try:
        #print(itemsCodes[i])
            if not itemsCodes[i] or pd.isna(itemsCodes[i]):
                itemsCodes.remove(itemsCodes[i])
            else:
                itemsCodes[i] = itemsCodes[i].replace(' ', '')
        except:
            break
    #print(itemsCodes)
    return orgName, items, itemsCodes

def extractServerDocInfo(serverPath):
    pdfs = glob.glob(serverPath + "/*.pdf")
    print(serverPath)
    # "реализация", "счет-фактура", "счет", "фактура",
    pdfs = list(filter(lambda p: any(name in p.replace('ё', 'е').lower() for name in ["упд", "универсальный"]), pdfs))
    updPdf = None
    if not pdfs:
        pdfs = glob.glob(serverPath + "/*.pdf")
        for pdf in pdfs:
            text, tables = get_info(pdf)
            for textLine in text:
                if "передаточный" in textLine.text.lower() and not "передача" in textLine.text.lower():
                    updPdf = pdf
                    break
        if not updPdf:
            return None, None, None, None, None, None, None
    else:
        updPdf = pdfs[0]

    print(updPdf)

    text, tables = get_info(updPdf)
    print('----')
    #print(text)
    if not text:
        text, tables = get_info(updPdf, rotated=True)
        os.remove("./rotated.pdf")
        if not text:
            return None, None, None, None, None, None, None
    cont = 0
    for textLine in text:
        if "передаточный" in textLine.text.lower():
            cont = 1
            break
    if not cont:
        text, tables = get_info(updPdf, rotated=True)
        os.remove("./rotated.pdf")
        if not text:
            return None, None, None, None, None, None, None
    updNum = ""
    updDate = ""
    updSumItems = ""
    updSumNDS = ""
    updSum = ""
    cont = 0
    for textLine in text:
        print(textLine.text.lower())
        if "передаточный" in textLine.text.lower():
            cont = 1
        if updDate == "":
            res = re.search(r"Счет-фактура\s№\s(\S*)\sот\s(\d{1,2}\s*\S+\s*\d\d\d\d)", textLine.text)
            if res:
                updNum, updDate = res.group(1), convDate(res.group(2))
        if updSum == "":
            res = re.search(r"всего к оплате\s*(?:(\(?\d?\)|)?)?\s*([\d ]+(?:,|.)\d\d)\s*.*?\s*([\d ]+(?:,|.)\d\d)?\s*([\d ]+(?:,|.)\d\d)", textLine.text.lower())
            if res:
                print(res)
                print(res.group(4))
                updSumItems = convSum(res.group(2))
                try:
                    updSumNDS = convSum(res.group(3))
                except:
                    updSumNDS = convSum(res.group(2))
                updSum = convSum(res.group(4))
    #updSum = ""
    print('9090990')
    print(updSum)
    print(updNum)
    if not cont:
        return None, None, None, None, None, None, None
    eps = 1
    while not updSum and eps < 16:
        text, tables = get_info(updPdf, eps = eps)
        for textLine in text:
            #print(textLine.text.lower())
            if updSum == "":
                res = re.search(
                    r"всего к оплате\s*(?:(\(?\d?\)|)?)?\s*([\d ]+(?:,|.)\d\d)\s*.*?\s*([\d ]+(?:,|.)\d\d)?\s*([\d ]+(?:,|.)\d\d)",
                    textLine.text.lower())
                if res:
                    updSumItems = convSum(res.group(2))
                    updSumNDS = convSum(res.group(3))
                    updSum = convSum(res.group(4))
        eps=eps+1
    print('9090990')
    print(updSum)
    print(updNum)
    gtdArr = []
    tables = camelot.read_pdf(filepath=updPdf, line_scale=100, flavor='lattice')
    if tables.n>2:
        #print('2')
        tables = camelot.read_pdf(filepath=updPdf, pages='all', line_scale=50, flavor='lattice')
    else:
        tables = camelot.read_pdf(filepath=updPdf, pages='all', line_scale=100,
                                  flavor='lattice')



    if tables.n > 1:
        # Если таблица в PDF размещена на нескольких страницах
        dfs = [table.df for table in tables]  # Собираем DataFrame из каждой таблицы


        if len(dfs[1].columns) > 1:
            if len(dfs[0].columns) != len(dfs[1].columns):
                dfs[0] = dfs[0].drop(0)
                dfs[0] = dfs[0].dropna(axis=1)
                dfs[0] = dfs[0].dropna(axis=0)
                for column in dfs[0].columns:
                    if dfs[0][column].nunique() == 1:
                        dfs[0] = dfs[0].drop(column, axis=1)
                dfs[0] = dfs[0].T.reset_index().T
                dfs[0] = dfs[0].reset_index()
                dfs[0] = dfs[0].drop(0)



            camelot_df = pd.concat(dfs, ignore_index=True)  # Объединяем DataFrame в один
        else:
            camelot_df = tables[0].df
    elif tables.n == 1:  # Обработка случая с одной таблицей
        camelot_df = tables[0].df
    camelot_df.to_excel('test.xlsx', index=False)

    ind = 0
    idx = 4
    for row in camelot_df.itertuples():
        ind += 1
        if 'страна' in str(row).lower():
            idx = ind

    #for row in camelot_df[idx:].itertuples():
    #    if row[16] != '—' and row[2].isdigit():
    #        gtdArr.append([int(row[2]), row[15], row[16].replace('\n', "")])

    is_first = True
    try:
        with pdfplumber.open(updPdf) as pdf:
            for page_num, page in enumerate(pdf.pages):
                # Find all tables on the current page
                tables = page.find_tables()
                if tables:
                    print(f"Found {len(tables)} tables on page {page_num + 1}")
                    for table_num, table in enumerate(tables):
                        # Extract the table data as a nested list
                        table_data = table.extract()
                        if table_data:
                            # Convert to a Pandas DataFrame for easier handling
                            pd.DataFrame(table_data, columns=table_data[0]).to_excel('test2.xlsx')

                            if is_first:
                                if "Универсальный" in table_data[0][0]:
                                    df = pd.DataFrame(table_data[1:], columns=table_data[0])
                                    df = df.drop(df.columns[0], axis=1)
                                else:
                                    df = pd.DataFrame(table_data[1:], columns=table_data[0])

                                #print(str(df.columns[0]))
                            else:
                                #df2 = pd.DataFrame(table_data, columns=table_data[0])
                                try:
                                    df = pd.concat([df, pd.DataFrame(table_data, columns=df.columns)], ignore_index=True)
                                except:
                                    try:
                                        df2 = pd.DataFrame(table_data[1:], columns=table_data[0])
                                        df2 = df2.drop(df2.columns[0], axis=1)
                                        df2.columns = df.columns
                                        df = pd.concat([df, df2], ignore_index=True)
                                    except:
                                        print("cringe table")
                            #print(df.columns)
                            is_first = False
                            #print(df)
                            if df.empty:
                                send_msg_to_bot(f'Для сделки {dealID} не найден корректный УПД. Пропускаю сделку')
                            else:
                                df.to_excel('test.xlsx')
    except:
        send_msg_to_bot(f'Для сделки {dealID} не найден корректный УПД. Пропускаю сделку')
        return None, None, None, None, None, None, None
    try:
        for row in df.itertuples():
            if row[2] and row[15]:
                if row[15] != '—' and "-" not in row[15] and '10а' not in row[15].replace('\n', "") and row[2].isdigit() and row[15] != '-' and "ДО ООО" not in row[16] and "—" not in row[16]:
                    if "/" in row[16]:
                        gtdArr.append([int(row[2]), row[15], row[16].replace('\n', ""), row[3].replace('\n', ""), row[7]])
                    else:
                        if row[17]:
                            if "/" in row[17]:
                                gtdArr.append([int(row[2]), row[16], row[17].replace('\n', ""), row[3].replace('\n', ""), row[7]])
    except:
        return None, None, None, None, None, None, None
    #print(len(df.columns))
    if not gtdArr and len(df.columns) < 16:
        df.insert(0, 'Supp_Col', "0")
        for row in df.itertuples():
            if row[2] and row[15]:
                if row[15] != '—' and "-" not in row[15] and '10а' not in row[15].replace('\n', "") and row[2].isdigit() and \
                        row[15] != '-' and "ДО ООО" not in row[16] and "—" not in row[16]:
                    if "/" in row[16]:
                        gtdArr.append([int(row[2]), row[15], row[16].replace('\n', ""), row[3].replace('\n', ""), row[7]])
                    else:
                        if row[17]:
                            if "/" in row[17]:
                                gtdArr.append(
                                    [int(row[2]), row[16], row[17].replace('\n', ""), row[3].replace('\n', ""), row[7]])

    print(gtdArr)

    #return 0
    return updNum, updDate, updSumItems, updSumNDS, updSum, gtdArr, df
