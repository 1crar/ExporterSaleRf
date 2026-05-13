"""1C document operations.

Function bodies are carried over from the legacy script so external behavior
stays the same while the module boundary becomes explicit.
"""

import datetime
import os
from pathlib import Path

import pandas as pd
import win32com.client
from thefuzz.fuzz import ratio

from ..bitrix import send_msg_to_bot
from ..parsing.utils import convDateTo1CFormat
from ..settings import SERVER_PATH as serverPath
from .printing import print_docs

def correctord(objSup, dfUpd, dfSum, connection):
    print('CORRECTION STARTED')
    objSup = objSup.GetObject()
    sum1 = 0
    found = 0
    #print(dfUpd)
    try:
        #for i in [6, 7, 8, 12]:
        #    dfUpd.iloc[:, i] = (
        #        dfUpd.iloc[:, i]
        #        .astype(str)
        #        .str.replace(',', '.', regex=False)
        #        .str.replace(' ', '', regex=False)
        #    )
        try:
            if pd.Index(dfUpd.columns.astype(str)).str.lower().str.contains('передаточный', na=False).any():
                dfUpd.columns = dfUpd.iloc[0]
                dfUpd = dfUpd.iloc[1:].reset_index(drop=True)
            dfUpd[dfUpd.columns[6]] = dfUpd[dfUpd.columns[6]].str.replace(',', '.').str.replace(' ', '')
            dfUpd[dfUpd.columns[7]] = dfUpd[dfUpd.columns[7]].str.replace(',', '.').str.replace(' ', '')
            dfUpd[dfUpd.columns[8]] = dfUpd[dfUpd.columns[8]].str.replace(',', '.').str.replace(' ', '')
            dfUpd[dfUpd.columns[12]] = dfUpd[dfUpd.columns[12]].str.replace(',', '.').str.replace(' ', '')
        except Exception as e:
            print(e)
            dfUpd.columns = dfUpd.iloc[0]
            dfUpd = dfUpd.iloc[1:].reset_index(drop=True)
            dfUpd.to_excel('testik.xlsx')
            dfUpd[dfUpd.columns[6]] = dfUpd[dfUpd.columns[6]].str.replace(',', '.').str.replace(' ', '')
            dfUpd[dfUpd.columns[7]] = dfUpd[dfUpd.columns[7]].str.replace(',', '.').str.replace(' ', '')
            dfUpd[dfUpd.columns[8]] = dfUpd[dfUpd.columns[8]].str.replace(',', '.').str.replace(' ', '')
            dfUpd[dfUpd.columns[12]] = dfUpd[dfUpd.columns[12]].str.replace(',', '.').str.replace(' ', '')

        sum1 = 0
        try:
            for row in dfUpd[2:].itertuples():
                if row[2] and row[2].isdigit():
                    if row[13] and not '1а' in row[3]:
                        sum1 = sum1 + float(row[13])
                        print(sum1)
            print('СУММА')
            print(sum1)
        except Exception as e:
            print('чет сумма не посичталась')
            print(e)
        arrTov = []
        for row in dfUpd[2:].itertuples():
            if row[3] and not 'наименование товара' in row[3].lower().replace('\n', " ") and not '1а' in row[3].lower().replace('\n', " "):
                found = 0
                for row2 in objSup.Ссылка.Товары:
                    if  row[3].lower().replace('\n', " ") == row2.Номенклатура.Наименование.lower():
                        print(row[3].replace('\n', " ") + " Найдено успешно")
                        print(row2.Номенклатура.Наименование)
                        arrTov.append(row2.Номенклатура)
                        found = 1
                        break
                    if connection.Справочники.Номенклатура.FindByDescription(row[3].replace('\n', " "), True).Наименование != '':
                        print(row[3].replace('\n', " ") + " Найдено успешно не в текущих документах")
                        print(connection.Справочники.Номенклатура.FindByDescription(row[3].replace('\n', " "), True).Наименование)
                        arrTov.append(connection.Справочники.Номенклатура.FindByDescription(row[3].replace('\n', " ")))
                        found = 1
                        break
                    if ratio(row[3].lower().replace('\n', " "), row2.Номенклатура.Наименование.lower()) > 95:
                        print(row[3].replace('\n', " ") + " Найдено успешно")
                        print(row2.Номенклатура.Наименование)
                        arrTov.append(row2.Номенклатура)
                        found = 1
                        break
                    if "поддон" in row[3].lower() or "упаковочная услуга" in row[3].lower() or "услуга упаковки" in row[3].lower() or "упаковка" in row[3].lower() or "бобин" in row[3].lower() or "паллет" in row[3].lower() or "тара" in row[3].lower() or "короб" in row[3].lower() or "барабан" in row[3].lower() or "фасовка" in row[3].lower() or "фасовочные" in row[3].lower() or "ящик" in row[3].lower() or "пленка" in row[3].lower() or "обрешетка" in row[3].lower():
                        print(row[3].replace('\n', " ") + " добавлено успешно")
                        dfUpd.iloc[row.Index, 3] = "Упаковка"
                        print(row[2])
                        print(row[3])
                        print(row[4])
                        arrTov.append(getattr(connection.Catalogs, "Номенклатура").FindByDescription("Упаковка"))
                        found = 1
                        break
                    if "самовывоз" in row[3].lower() or "доставка" in row[3].lower() or "транспортно-экспедиционные" in row[3].lower() or "ТЭУ" in row[3].lower() or "транспортные" in row[3].lower():
                        arrTov.append(getattr(connection.Catalogs, "Номенклатура").FindByDescription(
                            "Доставка (ТЭУ)"))
                        print(row[3].replace('\n', " ") + " добавлено успешно")
                        dfUpd.iloc[row.Index, 3] = "Доставка (ТЭУ)"
                        print(row[2])
                        print(row[3])
                        print(row[4])
                        found = 1
                        break
                if not found:
                    print(row[3].replace('\n', " ") + " не найдено")
                    break
    except Exception as e:
        print("что-то пошло не так с коррекшн")
        print(e)
    if float(sum1) == float(dfSum) and found == 1:
        print('Всё хорошо, корректирую на основании УПД')
        objSup.Товары.Clear()
        for row in dfUpd[2:].itertuples():
            if row[3] and not 'наименование товара' in row[3].lower().replace('\n', " ") and not '1а' in row[3].lower().replace('\n', " "):
                print(row[3].lower().replace('\n', " "))
                newline =  objSup.Товары.Add()
                for item in arrTov:
                    if row[3].lower().replace('\n', " ") == item.Наименование.lower():
                        newline.Номенклатура = item
                        break
                if newline.Номенклатура.Наименование == '':
                    for item in arrTov:
                        if ratio(row[3].lower().replace('\n', " "), item.Наименование.lower()) > 95:
                            newline.Номенклатура = item
                            break
                try:
                    newline.Цена = float(row[8])
                except:
                    print('Цена не указана')
                    newline.Цена = float(row[9])
                newline.Сумма = float(row[9])
                newline.СуммаСНДС = float(row[13])
                newline.СуммаНДС = float(row[13]) - float(row[9])
                try:
                    newline.Количество = float(row[7])
                    newline.КоличествоУпаковок = float(row[7])
                except:
                    print('Кол-во не указано')
                    newline.Количество = 1
                    newline.КоличествоУпаковок = 1


                if row[3] == 'Упаковка':
                    newline.СписатьНаРасходы = True
                    newline.СтатьяРасходов = getattr(connection.ChartsOfCharacteristicTypes,
                                                       "СтатьиРасходов").FindByDescription("Услуги по упаковке товара")
                if row[3] == 'Доставка (ТЭУ)':
                    newline.СписатьНаРасходы = True
                    newline.СтатьяРасходов = getattr(connection.ChartsOfCharacteristicTypes,
                                                       "СтатьиРасходов").FindByDescription("Доставка")
                print(row[11])
                if "22" in row[11]:
                    newline.СтавкаНДС = connection.Справочники.СтавкиНДС.FindByAttribute("ПеречислениеСтавкаНДС", connection.Перечисления.СтавкиНДС.НДС22)
                elif "20" in row[11]:
                    newline.СтавкаНДС = connection.Справочники.СтавкиНДС.FindByAttribute("ПеречислениеСтавкаНДС", connection.Перечисления.СтавкиНДС.НДС20)
                elif "18" in row[11]:
                    newline.СтавкаНДС = connection.Справочники.СтавкиНДС.FindByAttribute("ПеречислениеСтавкаНДС", connection.Перечисления.СтавкиНДС.НДС18)
                else:
                    newline.СтавкаНДС = connection.Справочники.СтавкиНДС.FindByAttribute("ПеречислениеСтавкаНДС", connection.Перечисления.СтавкиНДС.НДС0)
    objSup.ЖелаемаяДатаПоступления = connection.CurrentSessionDate()
    objSup.Write()
    print('correction ended')

def createPTU(connection, dosSs, dealID, RTUdate):
    docPTU = getattr(connection.Documents, "ПриобретениеТоваровУслуг").CreateDocument()

    query = connection.NewObject("Запрос")
    query.Text = """ВЫБРАТЬ
		КоммерческоеПредложениеКлиентуТовары.НомерСтроки КАК НомерСтроки,
		КоммерческоеПредложениеКлиентуТовары.Номенклатура КАК Номенклатура,
		КоммерческоеПредложениеКлиентуТовары.Характеристика КАК Характеристика,
		КоммерческоеПредложениеКлиентуТовары.ЕдиницаИзмерения КАК ЕдиницаИзмерения,
		СУММА(КоммерческоеПредложениеКлиентуТовары.Количество) КАК Количество,
		СУММА(КоммерческоеПредложениеКлиентуТовары.СуммаЭкспорт) КАК СуммаСНДС
	ПОМЕСТИТЬ КП
	ИЗ
		Документ.КоммерческоеПредложениеКлиенту.Товары КАК КоммерческоеПредложениеКлиентуТовары
	ГДЕ
		КоммерческоеПредложениеКлиентуТовары.Ссылка = &КП_Ссылка
		И КоммерческоеПредложениеКлиентуТовары.СуммаЭкспорт > 0
		И КоммерческоеПредложениеКлиентуТовары.Количество > 0
	
	СГРУППИРОВАТЬ ПО
		КоммерческоеПредложениеКлиентуТовары.НомерСтроки,
		КоммерческоеПредложениеКлиентуТовары.Номенклатура,
		КоммерческоеПредложениеКлиентуТовары.Характеристика,
		КоммерческоеПредложениеКлиентуТовары.ЕдиницаИзмерения
	;
	
	////////////////////////////////////////////////////////////////////////////////
	ВЫБРАТЬ
		ЗаказКлиентаТовары.КлючСвязи КАК КлючСвязи,
		ЗаказКлиентаТовары.Номенклатура КАК Номенклатура,
		ЗаказКлиентаТовары.Характеристика КАК Характеристика,
		СУММА(ЗаказКлиентаТовары.Количество) КАК Количество
	ПОМЕСТИТЬ ЗАКАЗ
	ИЗ
		Документ.ЗаказКлиента.Товары КАК ЗаказКлиентаТовары
	ГДЕ
		ЗаказКлиентаТовары.Ссылка = &ЗаказОснование
		
		
		И ЗаказКлиентаТовары.Э_ввд_ДоговорВнутреннейПерепродажи = &Э_ввд_ДоговорВнутреннейПерепродажи
	
	СГРУППИРОВАТЬ ПО
		ЗаказКлиентаТовары.КлючСвязи,
		ЗаказКлиентаТовары.Номенклатура,
		ЗаказКлиентаТовары.Характеристика
	;
	
	////////////////////////////////////////////////////////////////////////////////
	ВЫБРАТЬ
		РеализацияТоваровУслугТовары.КлючСвязи КАК КлючСвязи,
		РеализацияТоваровУслугТовары.Номенклатура КАК Номенклатура,
		РеализацияТоваровУслугТовары.Характеристика КАК Характеристика,
		СУММА(РеализацияТоваровУслугТовары.Количество) КАК Количество,
		СУММА(РеализацияТоваровУслугТовары.СуммаСНДС) КАК СуммаСНДС
	ПОМЕСТИТЬ РЕАЛИЗАЦИИ
	ИЗ
		Документ.РеализацияТоваровУслуг.Товары КАК РеализацияТоваровУслугТовары
	ГДЕ
		РеализацияТоваровУслугТовары.Ссылка.Проведен = ИСТИНА
		И РеализацияТоваровУслугТовары.Ссылка.Э_ввд_ЗаказОснование = &ЗаказОснование
		И РеализацияТоваровУслугТовары.Ссылка.Соглашение = &СоглашениеОрганизацииПродавца
	
	СГРУППИРОВАТЬ ПО
		РеализацияТоваровУслугТовары.КлючСвязи,
		РеализацияТоваровУслугТовары.Номенклатура,
		РеализацияТоваровУслугТовары.Характеристика
	;
	
	////////////////////////////////////////////////////////////////////////////////
	ВЫБРАТЬ
		КП.НомерСтроки КАК НомерСтроки,
		КП.Номенклатура КАК Номенклатура,
		КП.Характеристика КАК Характеристика,
		КП.ЕдиницаИзмерения КАК ЕдиницаИзмерения,
		ЗАКАЗ.Количество КАК Количество,
		КП.СуммаСНДС / КП.Количество * ЗАКАЗ.Количество КАК СуммаСНДС
	ПОМЕСТИТЬ КП_ЗАКАЗ
	ИЗ
		КП КАК КП
			ВНУТРЕННЕЕ СОЕДИНЕНИЕ ЗАКАЗ КАК ЗАКАЗ
			ПО КП.Номенклатура = ЗАКАЗ.Номенклатура
				И КП.Характеристика = ЗАКАЗ.Характеристика
				И КП.НомерСтроки = ЗАКАЗ.КлючСвязи 
	;
	
	////////////////////////////////////////////////////////////////////////////////
	ВЫБРАТЬ
		КП_ЗАКАЗ.НомерСтроки КАК НомерСтроки,
		КП_ЗАКАЗ.Номенклатура КАК Номенклатура,
		КП_ЗАКАЗ.Характеристика КАК Характеристика,
		КП_ЗАКАЗ.ЕдиницаИзмерения КАК ЕдиницаИзмерения,
		КП_ЗАКАЗ.Количество - ЕСТЬNULL(РЕАЛИЗАЦИИ.Количество, 0) КАК Количество,
		КП_ЗАКАЗ.СуммаСНДС - ЕСТЬNULL(РЕАЛИЗАЦИИ.СуммаСНДС, 0) КАК СуммаСНДС
	ПОМЕСТИТЬ ИТОГ
	ИЗ
		КП_ЗАКАЗ КАК КП_ЗАКАЗ
			ЛЕВОЕ СОЕДИНЕНИЕ РЕАЛИЗАЦИИ КАК РЕАЛИЗАЦИИ
			ПО КП_ЗАКАЗ.Номенклатура = РЕАЛИЗАЦИИ.Номенклатура
				И КП_ЗАКАЗ.Характеристика = РЕАЛИЗАЦИИ.Характеристика
				И КП_ЗАКАЗ.НомерСтроки = РЕАЛИЗАЦИИ.КлючСвязи
	;
	
	////////////////////////////////////////////////////////////////////////////////
	ВЫБРАТЬ
		ИТОГ.Номенклатура КАК Номенклатура,
		ИТОГ.Характеристика КАК Характеристика,
		ИТОГ.ЕдиницаИзмерения КАК ЕдиницаИзмерения,
		СУММА(ИТОГ.Количество) КАК Количество,
		СУММА(ИТОГ.Количество) КАК КоличествоУпаковок,
		СУММА(ИТОГ.СуммаСНДС) КАК СуммаСНДС
	ИЗ
		ИТОГ КАК ИТОГ
	ГДЕ
		ИТОГ.Количество <> 0
	
	СГРУППИРОВАТЬ ПО
		ИТОГ.НомерСтроки,
		ИТОГ.Номенклатура,
		ИТОГ.Характеристика,
		ИТОГ.ЕдиницаИзмерения
	
	УПОРЯДОЧИТЬ ПО
		ИТОГ.НомерСтроки;"""

    DP = dosSs.Товары[0].Э_ввд_ДоговорВнутреннейПерепродажи

    print('-----')
    print(DP.Наименование)

    query.SetParameter("ЗаказОснование", dosSs)
    query.SetParameter("КП_Ссылка", dosSs.ДокументОснование)
    query.SetParameter("СоглашениеОрганизацииПродавца", DP.СоглашениеОрганизацииПродавца)
    query.SetParameter("Э_ввд_ДоговорВнутреннейПерепродажи", DP)
    tablch = query.Execute().Unload()


    docPTU.Дата = connection.CurrentSessionDate()
    setattr(docPTU, "Дата", RTUDate + datetime.timedelta(hours=datetime.datetime.now().hour + 7) + datetime.timedelta(
        minutes=datetime.datetime.now().minute))
    docPTU.ХозяйственнаяОперация = connection.Перечисления.ХозяйственныеОперации.ЗакупкаУПоставщика
    docPTU.Организация = dosSs.Организация
    docPTU.Склад = dosSs.Склад
    docPTU.ИДСделкиБитрикс24 = dosSs.ИДСделкиБитрикс24
    docPTU.Э_ввд_ЗаказОснование = dosSs
    currency = connection.Константы.ВалютаРегламентированногоУчета.Get()
    docPTU.Валюта = currency
    docPTU.ВалютаВзаиморасчетов = currency
    docPTU.СпособДоставки = connection.Перечисления.СпособыДоставки.СиламиПоставщикаДоНашегоСклада
    docPTU.ЗакупкаПодДеятельность = connection.Перечисления.ТипыНалогообложенияНДС.ПродажаОблагаетсяНДС
    docPTU.НалогообложениеНДС = connection.Перечисления.ТипыНалогообложенияНДС.ПродажаОблагаетсяНДС
    docPTU.ЦенаВключаетНДС = True
    docPTU.ВариантПриемкиТоваров = connection.Перечисления.ВариантыПриемкиТоваров.РазделенаТолькоПоНакладным





    docPTU.Договор = DP.СоглашениеОрганизацииПокупателя
    docPTU.Контрагент = docPTU.Договор.Контрагент
    docPTU.Партнер = docPTU.Договор.Партнер




    docPTU.НомерВходящегоДокумента = dealID
    docPTU.ДатаВходящегоДокумента = connection.CurrentSessionDate()

    docPTU.Товары.Load(tablch)

    for strTovar in docPTU.Товары:
        strTovar.Склад = docPTU.Склад
        strTovar.Сумма = strTovar.СуммаСНДС
        strTovar.СтавкаНДС = connection.Справочники.СтавкиНДС.FindByAttribute("Ставка", 0)
        if strTovar.КоличествоУпаковок == 0:
            strTovar.Цена = strTovar.СуммаСНДС
        else:
            strTovar.Цена = strTovar.СуммаСНДС / strTovar.КоличествоУпаковок
    #docPTU.Номер = "КТУТ-004711"
    docPTU.Write()
    print(docPTU.Номер)
    docPTU.Write(connection.РежимЗаписиДокумента.Проведение)

    #ЗакупкиСервер.ЗаполнитьНомераГТДПоПоследнимПоступлениям(ДанныеФормыП.Товары, Ложь);
    return docPTU.Ссылка

def createRTU(connection, dosSs, dealID, RTUDate):
    docRTU = getattr(connection.Documents, "РеализацияТоваровУслуг").CreateDocument()

    query = connection.NewObject("Запрос")
    query.Text = """ВЫБРАТЬ
		КоммерческоеПредложениеКлиентуТовары.НомерСтроки КАК НомерСтроки,
		КоммерческоеПредложениеКлиентуТовары.Номенклатура КАК Номенклатура,
		КоммерческоеПредложениеКлиентуТовары.Характеристика КАК Характеристика,
		КоммерческоеПредложениеКлиентуТовары.ЕдиницаИзмерения КАК ЕдиницаИзмерения,
		СУММА(КоммерческоеПредложениеКлиентуТовары.Количество) КАК Количество,
		СУММА(КоммерческоеПредложениеКлиентуТовары.СуммаЭкспорт) КАК СуммаСНДС
	ПОМЕСТИТЬ КП
	ИЗ
		Документ.КоммерческоеПредложениеКлиенту.Товары КАК КоммерческоеПредложениеКлиентуТовары
	ГДЕ
		КоммерческоеПредложениеКлиентуТовары.Ссылка = &КП_Ссылка
		И КоммерческоеПредложениеКлиентуТовары.СуммаЭкспорт > 0
		И КоммерческоеПредложениеКлиентуТовары.Количество > 0

	СГРУППИРОВАТЬ ПО
		КоммерческоеПредложениеКлиентуТовары.НомерСтроки,
		КоммерческоеПредложениеКлиентуТовары.Номенклатура,
		КоммерческоеПредложениеКлиентуТовары.Характеристика,
		КоммерческоеПредложениеКлиентуТовары.ЕдиницаИзмерения
	;

	////////////////////////////////////////////////////////////////////////////////
	ВЫБРАТЬ
		ЗаказКлиентаТовары.КлючСвязи КАК КлючСвязи,
		ЗаказКлиентаТовары.Номенклатура КАК Номенклатура,
		ЗаказКлиентаТовары.Характеристика КАК Характеристика,
		СУММА(ЗаказКлиентаТовары.Количество) КАК Количество
	ПОМЕСТИТЬ ЗАКАЗ
	ИЗ
		Документ.ЗаказКлиента.Товары КАК ЗаказКлиентаТовары
	ГДЕ
		ЗаказКлиентаТовары.Ссылка = &ЗаказОснование
		И ЗаказКлиентаТовары.Э_ввд_ДоговорВнутреннейПерепродажи = &Э_ввд_ДоговорВнутреннейПерепродажи

	СГРУППИРОВАТЬ ПО
		ЗаказКлиентаТовары.КлючСвязи,
		ЗаказКлиентаТовары.Номенклатура,
		ЗаказКлиентаТовары.Характеристика
	;

	////////////////////////////////////////////////////////////////////////////////
	ВЫБРАТЬ
		РеализацияТоваровУслугТовары.КлючСвязи КАК КлючСвязи,
		РеализацияТоваровУслугТовары.Номенклатура КАК Номенклатура,
		РеализацияТоваровУслугТовары.Характеристика КАК Характеристика,
		СУММА(РеализацияТоваровУслугТовары.Количество) КАК Количество,
		СУММА(РеализацияТоваровУслугТовары.СуммаСНДС) КАК СуммаСНДС
	ПОМЕСТИТЬ РЕАЛИЗАЦИИ
	ИЗ
		Документ.РеализацияТоваровУслуг.Товары КАК РеализацияТоваровУслугТовары
	ГДЕ
		РеализацияТоваровУслугТовары.Ссылка.Проведен = ИСТИНА
		И РеализацияТоваровУслугТовары.Ссылка.Э_ввд_ЗаказОснование = &ЗаказОснование
		И РеализацияТоваровУслугТовары.Ссылка.Соглашение = &СоглашениеОрганизацииПродавца

	СГРУППИРОВАТЬ ПО
		РеализацияТоваровУслугТовары.КлючСвязи,
		РеализацияТоваровУслугТовары.Номенклатура,
		РеализацияТоваровУслугТовары.Характеристика
	;

	////////////////////////////////////////////////////////////////////////////////
	ВЫБРАТЬ
		КП.НомерСтроки КАК НомерСтроки,
		КП.Номенклатура КАК Номенклатура,
		КП.Характеристика КАК Характеристика,
		КП.ЕдиницаИзмерения КАК ЕдиницаИзмерения,
		ЗАКАЗ.Количество КАК Количество,
		КП.СуммаСНДС / КП.Количество * ЗАКАЗ.Количество КАК СуммаСНДС
	ПОМЕСТИТЬ КП_ЗАКАЗ
	ИЗ
		КП КАК КП
			ВНУТРЕННЕЕ СОЕДИНЕНИЕ ЗАКАЗ КАК ЗАКАЗ
			ПО КП.Номенклатура = ЗАКАЗ.Номенклатура
				И КП.Характеристика = ЗАКАЗ.Характеристика
				И КП.НомерСтроки = ЗАКАЗ.КлючСвязи 
	;

	////////////////////////////////////////////////////////////////////////////////
	ВЫБРАТЬ
		КП_ЗАКАЗ.НомерСтроки КАК НомерСтроки,
		КП_ЗАКАЗ.Номенклатура КАК Номенклатура,
		КП_ЗАКАЗ.Характеристика КАК Характеристика,
		КП_ЗАКАЗ.ЕдиницаИзмерения КАК ЕдиницаИзмерения,
		КП_ЗАКАЗ.Количество - ЕСТЬNULL(РЕАЛИЗАЦИИ.Количество, 0) КАК Количество,
		КП_ЗАКАЗ.СуммаСНДС - ЕСТЬNULL(РЕАЛИЗАЦИИ.СуммаСНДС, 0) КАК СуммаСНДС
	ПОМЕСТИТЬ ИТОГ
	ИЗ
		КП_ЗАКАЗ КАК КП_ЗАКАЗ
			ЛЕВОЕ СОЕДИНЕНИЕ РЕАЛИЗАЦИИ КАК РЕАЛИЗАЦИИ
			ПО КП_ЗАКАЗ.Номенклатура = РЕАЛИЗАЦИИ.Номенклатура
				И КП_ЗАКАЗ.Характеристика = РЕАЛИЗАЦИИ.Характеристика
				И КП_ЗАКАЗ.НомерСтроки = РЕАЛИЗАЦИИ.КлючСвязи
	;

	////////////////////////////////////////////////////////////////////////////////
	ВЫБРАТЬ
		ИТОГ.Номенклатура КАК Номенклатура,
		ИТОГ.Характеристика КАК Характеристика,
		ИТОГ.ЕдиницаИзмерения КАК ЕдиницаИзмерения,
		СУММА(ИТОГ.Количество) КАК Количество,
		СУММА(ИТОГ.Количество) КАК КоличествоУпаковок,
		СУММА(ИТОГ.СуммаСНДС) КАК СуммаСНДС
	ИЗ
		ИТОГ КАК ИТОГ
	ГДЕ
		ИТОГ.Количество <> 0

	СГРУППИРОВАТЬ ПО
		ИТОГ.НомерСтроки,
		ИТОГ.Номенклатура,
		ИТОГ.Характеристика,
		ИТОГ.ЕдиницаИзмерения

	УПОРЯДОЧИТЬ ПО
		ИТОГ.НомерСтроки;"""

    DP = dosSs.Товары[0].Э_ввд_ДоговорВнутреннейПерепродажи

    query.SetParameter("ЗаказОснование", dosSs)
    query.SetParameter("КП_Ссылка", dosSs.ДокументОснование)
    query.SetParameter("СоглашениеОрганизацииПродавца", DP.СоглашениеОрганизацииПродавца)
    query.SetParameter("Э_ввд_ДоговорВнутреннейПерепродажи", DP)
    tablch = query.Execute().Unload()

    try:
        docRTU.Дата = RTUDate
    except:
        docRTU.Дата = connection.CurrentSessionDate()
    if not docRTU.Дата:
        docRTU.Дата = connection.CurrentSessionDate()

    setattr(docRTU, "Дата", RTUDate + datetime.timedelta(hours=datetime.datetime.now().hour + 7) + datetime.timedelta(
        minutes=datetime.datetime.now().minute))

    print('----')
    print(RTUDate)
    print(docRTU.Дата)
    print(connection.CurrentSessionDate())
    docRTU.ХозяйственнаяОперация = connection.Перечисления.ХозяйственныеОперации.РеализацияКлиенту
    docRTU.Организация = dosSs.ДокументОснование.ОрганизацияЗакупа
    docRTU.Склад = dosSs.Склад
    docRTU.ИДСделкиБитрикс24 = dosSs.ИДСделкиБитрикс24
    docRTU.Э_ввд_ЗаказОснование = dosSs.Ссылка
    currency = connection.Константы.ВалютаРегламентированногоУчета.Get()
    docRTU.Валюта = currency
    docRTU.ВалютаВзаиморасчетов = currency
    docRTU.СпособДоставки = connection.Перечисления.СпособыДоставки.Самовывоз
    docRTU.ВариантОформленияПродажи = connection.Перечисления.ВариантыОформленияПродажи.РеализацияТоваровУслуг
    docRTU.ЗаказКлиента = connection.Документы.ЗаказКлиента.EmptyRef()
    docRTU.Согласован = True
    docRTU.Статус = connection.Перечисления.СтатусыРеализацийТоваровУслуг.Отгружено
    docRTU.НалогообложениеНДС = connection.Перечисления.ТипыНалогообложенияНДС.ПродажаНаЭкспорт
    docRTU.ЦенаВключаетНДС = True

    docRTU.Соглашение = DP.СоглашениеОрганизацииПродавца
    docRTU.Контрагент = docRTU.Соглашение.Контрагент
    docRTU.Партнер = docRTU.Соглашение.Партнер
    docRTU.ЦенаВключаетНДС = docRTU.Соглашение.ЦенаВключаетНДС



    docRTU.Товары.Load(tablch)
    for item in docRTU.Товары:
        #item.КлючСвязи = item.НомерСтроки

        item.СтавкаНДС = connection.Справочники.СтавкиНДС.FindByAttribute("ПеречислениеСтавкаНДС", connection.Перечисления.СтавкиНДС.НДС0)
        item.Склад = docRTU.Склад
        item.Сумма = item.СуммаСНДС
        if item.КоличествоУпаковок == 0:
            item.Цена = item.СуммаСНДС
        else:
            item.Цена = item.СуммаСНДС / item.КоличествоУпаковок
        item.КодТНВЭД = item.Номенклатура.КодТНВЭД


    docRTU.SetNewNumber()
    docRTU.Write()
    print(docRTU.Номер)
    docRTU.Write(connection.РежимЗаписиДокумента.Проведение)
    #docRTU.Write()
    return docRTU

def operations1C(dealID, updSum, updNum, updDate, connection, courierOrg, endCity, startCity, gtdArr, dfUpd, RTUdate, autoNum, has_miir_doc=True): #has_miir_doc _NEW

    # ---------------------------------
    # Проверка наличия в заказе клиента договора перепродажи
    # ----------------------------------
    # ==================================
    query = connection.NewObject("Запрос")
    query.Text = f"""
                     ВЫБРАТЬ
                         ЗаказКлиента.Ссылка КАК Ссылка
                     ИЗ
                         Документ.ЗаказКлиента КАК ЗаказКлиента
                     Где
                         ЗаказКлиента.Номер = &Номер
                  """

    query.SetParameter("Номер", str(dealID))
    selection = query.Execute().Choose()

    orderCorrect = False
    if selection.Next():
        if selection.Ссылка.Товары[0].Э_ввд_ДоговорВнутреннейПерепродажи.Наименование == "":
            pereprodazha = False
            print("Документ перепродажи не найден")
        else:
            pereprodazha = True
            print("Будет перепродажа")
        orderCorrect = True
    else:
        send_msg_to_bot(f"Документ с номером {dealID} не найден в 1С: УТ. Пропускаю сделку")
        return None
        raise ValueError(f"Документ с номером {dealID} не найден")
    # ==================================
        # Если МИИР РК без документа разрешения — допустимо только для сделок без перепродажи.
        # При перепродаже документ обязателен — останавливаем обработку.
    if not has_miir_doc and pereprodazha:
        send_msg_to_bot(f"Сделка {dealID}: МИИР РК, документ разрешения не вложен, но сделка с перепродажей. Требует ручной обработки.")
        print(f"Сделка {dealID}: МИИР РК без документа + перепродажа. Пропускаю.")
        return None

    updDate = convDateTo1CFormat(updDate)

    # ----------------------------------
    # Извлечение заказа поставщика и проверка совпадения суммы с УПД
    # ----------------------------------
    # ==================================
    querySupplier = connection.NewObject("Запрос")
    querySupplier.Text = f"""
                     ВЫБРАТЬ
                         ЗаказПоставщику.Ссылка КАК Ссылка,
                         ЗаказПоставщику.СуммаДокумента КАК СуммаДокумента,
                         ЗаказПоставщику.НалогообложениеНДС КАК НалогообложениеНДС,
                         ЗаказПоставщику.Товары КАК Товары
                     ИЗ
                         Документ.ЗаказПоставщику КАК ЗаказПоставщику
                     Где
                         ЗаказПоставщику.ИДСделкиБитрикс24 = &ИД
                  """

    querySupplier.SetParameter("ИД", str(dealID))
    selectionSupplier = querySupplier.Execute().Choose()
    #selectionSupplierEx = False

    queryProposal = connection.NewObject("Запрос")
    queryProposal.Text = f"""
                                             ВЫБРАТЬ
                                                 КоммерческоеПредложениеКлиенту.Ссылка КАК Ссылка
                                             ИЗ
                                                 Документ.КоммерческоеПредложениеКлиенту КАК КоммерческоеПредложениеКлиенту
                                             Где
                                                 КоммерческоеПредложениеКлиенту.ИДСделкиБитрикс24 = &ИД
                                          """

    queryProposal.SetParameter("ИД", str(dealID))
    selectionProposal = queryProposal.Execute().Choose()

    if not selectionProposal.Next():
        print(f"КоммерческоеПредложениеКлиенту по номеру {dealID} не найденo")
        return
    if pereprodazha:
        if selectionProposal.Ссылка.ОрганизацияЗакупа.Наименование != selection.Ссылка.Товары[0].Э_ввд_ДоговорВнутреннейПерепродажи.ОрганизацияПродавец.Наименование:
            print(selectionProposal.Ссылка.ОрганизацияЗакупа.Наименование)
            print(selection.Ссылка.Товары[0].Э_ввд_ДоговорВнутреннейПерепродажи.ОрганизацияПродавец.Наименование)
            send_msg_to_bot(f"Не совпадают организации в КП и Заказе клиента. Пропускаю сделку {dealID}")
            print(f"Не совпадают организации в КП и Заказе клиента. Пропускаю сделку {dealID}")
            return None
    if not selectionSupplier.Next():
        # raise ValueError(f"ЗаказПоставщику по номеру {dealID} не найдено")
        print(f"ЗаказПоставщику по номеру {dealID} не найдено")
        objSupplier = getattr(connection.Documents, "ЗаказПоставщику").CreateDocument()
        setattr(objSupplier, "Дата", connection.CurrentSessionDate())


        post = selectionProposal.Ссылка.Товары[0].Поставщик
        print(post.Наименование)
        for stroka in selectionProposal.Ссылка.Товары:
            if post.Наименование == "":
                post = stroka.Поставщик
            if post.Наименование != "" and stroka.Поставщик != post and stroka.Поставщик.Наименование != "":
                print("Много поставщиков! Пропускаю")
                return

        objSupplier.Fill(selection.Ссылка)
        objSupplier.Партнер = post.Партнер
        objSupplier.Контрагент = post
        objSupplier.ДокументОснование = selection.Ссылка
        objSupplier.Склад = selection.Ссылка.Склад

        print(objSupplier.Партнер.Наименование)
        print(objSupplier.Контрагент.Наименование)


        #objSupplier.Партнер = post.ОсновнойДоговор
        objSupplier.Write()

        dfUpd[dfUpd.columns[7]] = dfUpd[dfUpd.columns[7]].str.replace(',', '.').str.replace(' ', '')
        dfUpd[dfUpd.columns[8]] = dfUpd[dfUpd.columns[8]].str.replace(',', '.').str.replace(' ', '')
        dfUpd[dfUpd.columns[12]] = dfUpd[dfUpd.columns[12]].str.replace(',', '.').str.replace(' ', '')


        arrTov = []
        for poructStr in objSupplier.Ссылка.Товары:
            if not "услуга 20%" in poructStr.Номенклатура.Наименование.lower():
                arrTov.append(poructStr.Наименование.lower())

        for row in dfUpd[2:].itertuples():
            if row[2] and row[2].isdigit():
                if row[3].lower() in arrTov:
                    continue
                for poructStr in objSupplier.Ссылка.Товары:
                    if "услуга 20%" in poructStr.Номенклатура.Наименование.lower():
                        if "самовывоз" in row[3].lower():
                            poructStr.Номенклатура = getattr(connection.Catalogs, "Номенклатура").FindByDescription("Доставка (ТЭУ)")
                            poructStr.СписатьНаРасходы = True
                            poructStr.СтатьяРасходов = getattr(connection.ChartsOfCharacteristicTypes, "СтатьиРасходов").FindByDescription("Доставка")
                        if "поддон" in row[3].lower():
                            poructStr.Номенклатура = getattr(connection.Catalogs, "Номенклатура").FindByDescription("Упаковка")
                            poructStr.СписатьНаРасходы = True
                            poructStr.СтатьяРасходов = getattr(connection.ChartsOfCharacteristicTypes, "СтатьиРасходов").FindByDescription("Услуги по упаковке товара")
                        break

        objSupplier.Write(connection.РежимЗаписиДокумента.Проведение)
        selectionSupplier = objSupplier.Ссылка
        nonNDSsum = selectionSupplier.СуммаДокумента
    else:
        nonNDSsum = selectionSupplier.СуммаДокумента



    clientServerPath = None

    if orderCorrect:
        if float(updSum) != nonNDSsum:
            print('NOT EQ')

            # ==================================


            # -----------------------------
            # Обновление заказа поставщика ---
            # -----------------------------
            # ==================================

            objSupplier = selectionSupplier.Ссылка.GetObject()

            objSupplier.Товары.Clear()
            '''for el in objSupplier.Metadata().ТабличныеЧасти.Товары.Реквизиты:
                print(el.Имя)
            print("----------------------------")
            for el in selectionProposal.Ссылка.GetObject().Metadata().ТабличныеЧасти.Товары.Реквизиты:
                print(el.Имя)'''
            for poructStr in selectionProposal.Ссылка.Товары:
                if "наценка" in poructStr.Номенклатура.Наименование.lower() or "транспортные расходы" in poructStr.Номенклатура.Наименование.lower():
                    continue
                newstr = objSupplier.Товары.Add()
                #print("Количество продукта:", poructStr.Количество)
                newstr.Номенклатура = poructStr.Номенклатура
                newstr.Количество = poructStr.Количество
                newstr.КоличествоУпаковок = poructStr.Количество
                newstr.Цена = poructStr.ЦенаПоставщикаВРублях
                newstr.СтавкаНДС = poructStr.СтавкаНДСПоставщика
                newstr.СуммаНДС = poructStr.СуммаНДСПоставщика
                newstr.Сумма = poructStr.СтоимостьПоставщика
                newstr.СуммаСНДС = newstr.СуммаНДС + newstr.Сумма
            objSupplier.Fill(selectionProposal.Ссылка)

            arrTov = []
            for poructStr in objSupplier.Ссылка.Товары:
                if not "услуга 20%" in poructStr.Номенклатура.Наименование.lower():
                    arrTov.append(poructStr.Номенклатура.Наименование.lower())

            for row in dfUpd[2:].itertuples():
                print(row)
                if row[2] and row[2].isdigit():
                    if row[3].lower() in arrTov:
                        continue
                    for poructStr in objSupplier.Товары:
                        if "услуга 20%" in poructStr.Номенклатура.Наименование.lower():
                            if "самовывоз" in row[3].lower():
                                poructStr.Номенклатура = getattr(connection.Catalogs, "Номенклатура").FindByDescription(
                                    "Доставка (ТЭУ)")
                                poructStr.СписатьНаРасходы = True
                                poructStr.СтатьяРасходов = getattr(connection.ChartsOfCharacteristicTypes,
                                                                   "СтатьиРасходов").FindByDescription("Доставка")
                            if "поддон" in row[3].lower():
                                poructStr.Номенклатура = getattr(connection.Catalogs, "Номенклатура").FindByDescription(
                                    "Упаковка")
                                poructStr.СписатьНаРасходы = True
                                poructStr.СтатьяРасходов = getattr(connection.ChartsOfCharacteristicTypes,
                                                                   "СтатьиРасходов").FindByDescription(
                                    "Услуги по упаковке товара")
                            break
            objSupplier.Склад = selection.Ссылка.Склад
            objSupplier.Write()
            # ==================================

            print(dfUpd.iloc[:, 7])
            dfUpd.iloc[:, 7] = dfUpd.iloc[:, 7].astype(str).str.replace(',', '.').str.replace(' ', '')
            dfUpd.iloc[:, 8] = dfUpd.iloc[:, 8].astype(str).str.replace(',', '.').str.replace(' ', '')
            dfUpd.iloc[:, 12] = dfUpd.iloc[:, 12].astype(str).str.replace(',', '.').str.replace(' ', '')
            print(updSum)
            print(objSupplier.СуммаДокумента)
            try:
                if float(objSupplier.СуммаДокумента) != float(updSum):
                    for row in dfUpd[2:].itertuples():
                        if row[2] and row[2].isdigit():
                            strNom = int(row[2])-1
                            print(row)
                            if float(row[8]) != float(objSupplier.Товары[strNom].Цена) or float(row[9]) != float(objSupplier.Товары[strNom].Сумма) or float(row[13]) != float(objSupplier.Товары[strNom].СуммаСНДС):
                                objSupplier.Товары[strNom].Цена = float(row[8])
                                objSupplier.Товары[strNom].Сумма = float(row[9])
                                objSupplier.Товары[strNom].СуммаСНДС = float(row[13])
                                objSupplier.Товары[strNom].СуммаНДС = float(row[13]) - float(row[9])
            except Exception as e:
                print(e)
                print("Базовый корректор не прошел")
            objSupplier.Write()
            #objSupplier.Write(connection.РежимЗаписиДокумента.Проведение)

        correctord(selectionSupplier.Ссылка, dfUpd, updSum, connection)
        # ----------------------------------
        # Создание приобретения товаров и услуг на основе заказа поставщика
        # ----------------------------------
        # ==================================

        #objSupplier = selectionSupplier.Ссылка.GetObject()
        #objSupplier.Write(connection.РежимЗаписиДокумента.Проведение)
        creatorAcqDoc = getattr(connection.Documents, "ПриобретениеТоваровУслуг")
        orgName = selection.Ссылка.Организация.Наименование
        comm = f"{orgName} / {dealID}"
        if not pereprodazha:
            try:
                comm = f"{selection.Ссылка.Контрагент.Наименование} / {dealID}"
            except Exception as e:
                print("Комментарий не ввел")
        doc = creatorAcqDoc.CreateDocument()
        setattr(doc, "НомерВходящегоДокумента", updNum)
        setattr(doc, "Дата", updDate + datetime.timedelta(hours = datetime.datetime.now().hour) + datetime.timedelta(minutes = datetime.datetime.now().minute))
        setattr(doc, "Комментарий", comm)
        setattr(doc, "ДатаВходящегоДокумента", updDate)
        setattr(doc, "НаименованиеВходящегоДокумента", "УПД")

        doc.Fill(selectionSupplier.Ссылка)
        setattr(doc, "Склад", selectionSupplier.Ссылка.Склад)
        #setattr(doc, "Номер", "РЭУТ-093696")
        doc.Товары.Clear()
        for poructStr in selectionSupplier.Ссылка.Товары:
            newstr = doc.Товары.Add()
            print(poructStr.Номенклатура.Наименование)
            newstr.Номенклатура = poructStr.Номенклатура
            newstr.Количество = poructStr.Количество
            newstr.КоличествоУпаковок = poructStr.Количество
            newstr.Цена = poructStr.Цена
            newstr.СтавкаНДС = poructStr.СтавкаНДС
            newstr.СуммаНДС = poructStr.СуммаНДС
            newstr.Сумма = poructStr.Сумма
            newstr.СуммаСНДС = poructStr.СуммаСНДС
            newstr.КодСтроки = poructStr.КодСтроки
            newstr.Склад = selectionSupplier.Ссылка.Склад

        for item in gtdArr:

            doc.Товары[item[0]-1].НомерГТД = getattr(connection.Catalogs, "НомераГТД").FindByAttribute("РегистрационныйНомер", str(item[2]))
            print('-----')
            print(getattr(connection.Catalogs, "НомераГТД").FindByAttribute("РегистрационныйНомер", str(item[2])).Код)
            print(doc.Товары[item[0]-1].НомерГТД.Код)
            print('--------')

            if doc.Товары[item[0]-1].НомерГТД.РегистрационныйНомер == '':
                print('34')
                new_gtd = getattr(connection.Catalogs, "НомераГТД").CreateItem()
                #new_gtd.Наименование = str(item[2])
                new_gtd.Код = str(item[2])
                new_gtd.ТипНомераГТД = connection.Перечисления.ТипыНомеровГТД.НомерГТД
                new_gtd.РегистрационныйНомер = str(item[2])
                if not "тайвань" in str(item[1]).lower():
                    new_gtd.СтранаПроисхождения = getattr(connection.Catalogs, "СтраныМира").FindByDescription(
                        str(item[1]).upper())
                else:
                    new_gtd.СтранаПроисхождения = getattr(connection.Catalogs, "СтраныМира").FindByDescription("ТАЙВАНЬ (КИТАЙ)")
                print(getattr(connection.Catalogs, "СтраныМира").FindByDescription(str(item[1]).upper()).Наименование)
                new_gtd.Write()
                doc.Товары[item[0] - 1].НомерГТД = new_gtd
                print(doc.Товары[item[0] - 1].НомерГТД.РегистрационныйНомер)
                doc.Товары[item[0] - 1].НомерГТД = getattr(connection.Catalogs, "НомераГТД").FindByAttribute(
                    "РегистрационныйНомер", str(item[2]))
                print(doc.Товары[item[0] - 1].НомерГТД.РегистрационныйНомер)
        doc.Write()
        try:
            doc.Write(connection.РежимЗаписиДокумента.Проведение)
        except Exception as e:
            print("Не смог провести ПТУ!!!")
            print(e)
        print("PTU " + doc.Номер)
        createInvoice(doc.Ссылка, updDate, updNum, connection)



        if pereprodazha:

            docPTU = createPTU(connection, selection.Ссылка, dealID, RTUDate)

            # createInvoice(docPTU, updDate, updNum, connection)

            docRTU = createRTU(connection, selection.Ссылка, dealID, RTUDate)

            creatorInvoiceDoc = getattr(connection.Documents, "СчетФактураВыданный")
            doc = creatorInvoiceDoc.CreateDocument()
            doc.Fill(docRTU.Ссылка)
            doc.ДатаВыставления = connection.CurrentSessionDate()
            doc.Write(connection.РежимЗаписиДокумента.Проведение)
            print(doc.Номер)

            creatorCMR = getattr(connection.Documents, "ООО_CMR")
            cmrDoc = creatorCMR.CreateDocument()
            cmrDoc.Fill(docRTU.Ссылка)
            setattr(cmrDoc, "Дата", connection.CurrentSessionDate())
            setattr(cmrDoc, "Номер", docRTU.Ссылка.Номер[2:])
            setattr(cmrDoc, "Грузоперевозчик", courierOrg)
            setattr(cmrDoc, "ГородПогрузки", endCity)
            setattr(cmrDoc, "ГородРазгрузки", startCity)
            setattr(cmrDoc, "СтранаРазгрузки", getattr(cmrDoc, "Грузополучатель").СтранаРегистрации.Наименование)
            try:
                print(autoNum)
                setattr(cmrDoc, "Автомобиль", autoNum[0].split('/')[0])
                setattr(cmrDoc, "Прицеп", autoNum[0].split('/')[1])
            except:
                print("Нет номера авто")

            cmrDoc.Write()
            orgName = orgName.upper()
            exporterOrg = selectionSupplier.Ссылка.Организация.Наименование
            print('123456')
            print(exporterOrg)
            print(orgName)
            isCourierPartner = ratio(exporterOrg, "ВЭД ПАРТНЕР ООО") > 90
            isCourierExporter = ratio(exporterOrg, "ЭКСПОРТЕР ООО") > 70
            isCourierFastResheniya = ratio(exporterOrg, "БЫСТРЫЕ РЕШЕНИЯ") > 70

            isTrial = ratio(orgName, "ТРИАЛ-ПВ, ТОО") > 90
            isPartner = ratio(orgName, "ВЭД ПАРТНЕР ТОО") > 90
            isPartnerAlmati = ratio(orgName, "ВЭД Партнер-Алматы ТОО") > 90
            isPartnerOsOO = ratio(orgName, "ВЭД ПАРТНЕР, ОСОО") > 90

            if isTrial:
                if isCourierPartner:
                    clientServerPath = serverPath + "ООО ВЭД ПАРТНЕР\\ПОКУПАТЕЛИ\\Т\\ТРИАЛ-ПВ, ТОО"
                elif isCourierExporter:
                    clientServerPath = serverPath + "ЭКСПОРТЕР НОВЫЙ\\КЛИЕНТЫ Эн\\КАЗАХСТАН\\Т\\ТРИАЛ-ПВ"
                elif isCourierFastResheniya:
                    clientServerPath = serverPath + "БЫСТРЫЕ РЕШЕНИЯ\\Клиенты\\Т\\ТРИАЛ-ПВ, ТОО"
            elif isPartner:
                if isCourierPartner:
                    clientServerPath = serverPath + "ООО ВЭД ПАРТНЕР\\ПОКУПАТЕЛИ\\В\\ВЭД Партнер, ТОО"
                elif isCourierExporter:
                    clientServerPath = serverPath + "ЭКСПОРТЕР НОВЫЙ\\КЛИЕНТЫ Эн\\КАЗАХСТАН\\В\\ВЭД ПАРТНЕР ТОО\\Отгрузочные документы по сделкам"
                elif isCourierFastResheniya:
                    clientServerPath = serverPath + "БЫСТРЫЕ РЕШЕНИЯ\\Клиенты\\В\\ВЭД Партнер, ТОО"
            elif isPartnerAlmati:
                if isCourierPartner:
                    clientServerPath = serverPath + "ООО ВЭД ПАРТНЕР\\ПОКУПАТЕЛИ\\В\\ВЭД Партнер-Алматы ТОО"
                elif isCourierExporter:
                    clientServerPath = serverPath + "ЭКСПОРТЕР НОВЫЙ\\КЛИЕНТЫ Эн\\КАЗАХСТАН\\В\\ВЭД Партнер-Алматы ТОО\\Документы для отгрузки"
                elif isCourierFastResheniya:
                    clientServerPath = serverPath + "БЫСТРЫЕ РЕШЕНИЯ\\Клиенты\\ВЭД-Партнер Алматы, ТОО"
            elif isPartnerOsOO:
                if isCourierPartner:
                    clientServerPath = serverPath + "ООО ВЭД ПАРТНЕР\\ПОКУПАТЕЛИ\\В\\ВЭД Партнер, ОсОО"
                elif isCourierExporter:
                    clientServerPath = serverPath + "ЭКСПОРТЕР НОВЫЙ\\КЛИЕНТЫ Эн\\КИРГИЗИЯ\\В\\ВЭД ОсОО"
                elif isCourierFastResheniya:
                    clientServerPath = serverPath + "БЫСТРЫЕ РЕШЕНИЯ\\Клиенты\\ВЭД Партнер ОсОО"
            if not clientServerPath:
                print("ПРЕДУПРЕЖДЕНИЕ: Неизвестный доставщик или организация! Формируется отдельная папка для документов")
                clientServerPath = f"{serverPath}\\{exporterOrg.upper()}\\{orgName.upper()[0]}\\{orgName.upper()}\\{dealID}"
            else:
                clientServerPath += f"\\{dealID}"
            os.makedirs(clientServerPath, exist_ok=True)
            print_docs(cmrDoc, doc, clientServerPath, connection)
        else:
            print('---------')
            objOrder = selection.Ссылка.GetObject()

            for poructStr in objOrder.Товары:
                poructStr.ВариантОбеспечения = connection.Перечисления.ВариантыОбеспечения.Отгрузить

            objOrder.Write()
            objOrder.Write(connection.РежимЗаписиДокумента.Проведение)
            creatorAcqDoc = getattr(connection.Documents, "РеализацияТоваровУслуг")
            docRTU = creatorAcqDoc.CreateDocument()
            docRTU.Fill(selection.Ссылка)
            docRTU.Товары.Clear()
            for poructStr in selection.Ссылка.Товары:
                newstr = docRTU.Товары.Add()
                # print(poructStr.Номенклатура.Наименование)
                newstr.Номенклатура = poructStr.Номенклатура
                newstr.Количество = poructStr.Количество
                newstr.КоличествоУпаковок = poructStr.Количество
                newstr.Цена = poructStr.Цена
                newstr.СтавкаНДС = poructStr.СтавкаНДС
                newstr.СуммаНДС = poructStr.СуммаНДС
                newstr.Сумма = poructStr.Сумма
                newstr.СуммаСНДС = poructStr.СуммаСНДС
                newstr.КодСтроки = poructStr.КодСтроки
                newstr.Склад = selectionSupplier.Ссылка.Склад
            setattr(docRTU, "Дата", connection.CurrentSessionDate())
            try:
                docRTU.НалогообложениеНДС = connection.Перечисления.ТипыНалогообложенияНДС.ПродажаНаЭкспорт
            except Exception as e:
                print(e)
            docRTU.Write()
            try:
                for poructStr in docRTU.Товары:
                    poructStr.КодТНВЭД = poructStr.Номенклатура.КодТНВЭД


                #vn_obr = connection.ВнешниеОбработки.Create(fr"Print.epf", False)
                #vn_obr.TNVED(docRTU.Ссылка)
            except Exception as e:
                print(e)
                print("ТН ВЭД не проставил")
            try:
                docRTU.Write(connection.РежимЗаписиДокумента.Проведение)
            except Exception as e:
                send_msg_to_bot(f"Не смог провести РТУ для сделки {dealID}. Проверьте настройки 1С")
                print(f"Не смог провести РТУ для сделки {dealID}. Проверьте настройки 1С")
                print(e)
            creatorInvoiceDoc = getattr(connection.Documents, "СчетФактураВыданный")
            doc = creatorInvoiceDoc.CreateDocument()
            doc.Fill(docRTU.Ссылка)
            doc.ДатаВыставления = connection.CurrentSessionDate()
            # doc.Write()
            doc.Write(connection.РежимЗаписиДокумента.Проведение)
            print(doc.Номер)

            creatorCMR = getattr(connection.Documents, "ООО_CMR")
            cmrDoc = creatorCMR.CreateDocument()
            cmrDoc.Fill(docRTU.Ссылка)
            setattr(cmrDoc, "Дата", connection.CurrentSessionDate())
            setattr(cmrDoc, "Номер", docRTU.Ссылка.Номер[2:])
            setattr(cmrDoc, "Грузоперевозчик", courierOrg)
            setattr(cmrDoc, "ГородПогрузки", endCity)
            setattr(cmrDoc, "ГородРазгрузки", startCity)
            setattr(cmrDoc, "СтранаРазгрузки", getattr(cmrDoc, "Грузополучатель").СтранаРегистрации.Наименование)
            try:
                print(autoNum)
                setattr(cmrDoc, "Автомобиль", autoNum[0].split('/')[0])
                setattr(cmrDoc, "Прицеп", autoNum[0].split('/')[1])
            except:
                print("Нет номера авто")

            cmrDoc.Write()
            orgName = orgName.upper()
            exporterOrg = selectionSupplier.Ссылка.Организация.Наименование
            print('123456')
            print(exporterOrg)
            print(orgName)
            isCourierPartner = ratio(exporterOrg, "ВЭД ПАРТНЕР ООО") > 90
            isCourierExporter = ratio(exporterOrg, "ЭКСПОРТЕР") > 70
            isCourierFastResheniya = ratio(exporterOrg, "БЫСТРЫЕ РЕШЕНИЯ") > 70
            isCourierTorgExp = ratio(exporterOrg, "ТОРГЭКСПЕРТ") > 70

            if isCourierPartner:
                clientServerPath = serverPath + "ООО ВЭД ПАРТНЕР"
            elif isCourierFastResheniya:
                clientServerPath = serverPath + "БЫСТРЫЕ РЕШЕНИЯ\\Клиенты"
            elif isCourierExporter:
                    clientServerPath = serverPath + "ЭКСПОРТЕР НОВЫЙ\\КЛИЕНТЫ Эн"
            elif isCourierTorgExp:
                    clientServerPath = serverPath + "БЫСТРЫЕ РЕШЕНИЯ\\Клиенты\\ВЭД Партнер ОсОО"
            if not clientServerPath:
                print(
                    "ПРЕДУПРЕЖДЕНИЕ: Неизвестная организация! Пропускаю вложение документов")
                clientServerPath = r'C:\Users\dev-miir-rk\PycharmProjects\AutoDocAccounting\temp'
            else:
                print(clientServerPath)
                if isCourierPartner:
                    results = [p for p in Path(clientServerPath+"\\КЛИЕНТЫ").rglob(dealID) if p.is_dir()]
                    if not results:
                        results = [p for p in Path(clientServerPath + "\\ПОКУПАТЕЛИ").rglob(dealID) if p.is_dir()]
                else:
                    results = [p for p in Path(clientServerPath).rglob(dealID) if p.is_dir()]
                try:
                    clientServerPath = str(results[0])
                except:
                    try:
                        if isCourierPartner:
                            results = [p for p in Path(clientServerPath + "\\КЛИЕНТЫ").rglob("спец*" + dealID) if p.is_dir()]
                            if not results:
                                results = [p for p in Path(clientServerPath + "\\ПОКУПАТЕЛИ").rglob("спец*" + dealID) if
                                           p.is_dir()]
                        else:
                            results = [p for p in Path(clientServerPath).rglob("спец*" + dealID) if p.is_dir()]
                        clientServerPath = str(results[0])
                    except:
                        clientServerPath = r'C:\Users\dev-miir-rk\PycharmProjects\AutoDocAccounting\temp'
            print_docs(cmrDoc, doc, clientServerPath, connection)

            print('---------')

        # ==================================
    else:
        send_msg_to_bot(f"Документ с номером {dealID} не найден")
        '''creatorAcqDoc = getattr(connection.Documents, "ПриобретениеТоваровУслуг")
        doc = creatorAcqDoc.CreateDocument()
        doc.Fill(selectionSupplier.Ссылка)

        setattr(doc, "НомерВходящегоДокумента", updNum)
        setattr(doc, "Дата", updDate)
        setattr(doc, "ДатаВходящегоДокумента", updDate)
        setattr(doc, "НаименованиеВходящегоДокумента", "УПД")
        doc.write()
        num = getattr(doc, "Номер")
        createInvoice(num, updDate, connection)

        for item in selection.Товары:
            print(item.Действия)'''

    return clientServerPath

def createInvoice(docSs, updDate, updNum, connection):
    creatorInvoiceDoc = getattr(connection.Documents, "СчетФактураПолученный")
    doc = creatorInvoiceDoc.CreateDocument()
    doc.Fill(True)

    setattr(doc, "Организация", docSs.Организация)
    setattr(doc, "Контрагент", docSs.Контрагент)
    setattr(doc, "Валюта", docSs.Договор.ВалютаВзаиморасчетов)
    if docSs.Договор.ВалютаВзаиморасчетов.Наименование == '':
        setattr(doc, "Валюта", docSs.Валюта)
    setattr(doc, "ДатаСоставления", updDate)
    setattr(doc, "Номер", updNum)
    setattr(doc, "ИннКонтрагента", docSs.Контрагент.ИНН)
    setattr(doc, "КППКонтрагента", docSs.Контрагент.КПП)

    newstr = doc.ДокументыОснования.Add()
    newstr.ДокументОснование = docSs
    isOneNDS = True
    SumNds = 0
    lastNds = docSs.Товары[0].СтавкаНДС
    for item in docSs.Товары:
        if item.СтавкаНДС != lastNds:
            isOneNDS = False
        SumNds = SumNds + item.СуммаНдс
    if isOneNDS:
        newstr.СтавкаНДС = lastNds
    newstr.Сумма = docSs.СуммаДокумента
    newstr.СуммаНДС = SumNds

    setattr(doc, "Дата", updDate)

    doc.Write(connection.РежимЗаписиДокумента.Проведение)

def check1Cdocdeal(dealID, orgname, itemcodes, connection):

    # ---------------------------------
    # Сверка инфо в заказе клиента с разрешительным документом в битриксе
    # ----------------------------------
    # ==================================
    query = connection.NewObject("Запрос")
    query.Text = f"""
                     ВЫБРАТЬ
                         ЗаказКлиента.Ссылка КАК Ссылка,
                         ЗаказКлиента.Организация КАК Организация
                     ИЗ
                         Документ.ЗаказКлиента КАК ЗаказКлиента
                     Где
                         ЗаказКлиента.Номер = &Номер
                  """

    query.SetParameter("Номер", str(dealID))
    selection = query.Execute().Choose()


    if selection.Next():
        if selection.Организация.Наименование != orgname:
            return False
        counter = 0
        for stritems in selection.Ссылка.Товары:
            if stritems.Номенклатура.КодТНВЭД.Код != itemcodes[counter]:
                counter += 1
                return False
    else:
        return False
        raise ValueError(f"Документ с номером {dealID} не найден")

    return True

def testupdprint():
    connection = win32com.client.Dispatch("V83.COMConnector").Connect('Srvr="1c-dev";Ref="logist_2025_rbc";'
                                                                         'Usr="РарусRPA";Pwd="7zevecoC";')
    query = connection.NewObject("Запрос")
    query.Text = f"""
                         ВЫБРАТЬ
                        РеализацияТоваровУслуг.Ссылка КАК Ссылка
                    ИЗ
                        Документ.РеализацияТоваровУслуг КАК РеализацияТоваровУслуг
                    ГДЕ
	                    РеализацияТоваровУслуг.Номер = &Номер
                      """

    query.SetParameter("Номер", "ТЭУТ-000001")
    selection = query.Execute().Choose()
    if selection.Next():
        print("1")
        print(selection)
    vn_obr = connection.ВнешниеОбработки.Create(fr"Print.epf", False)
    print(selection.Ссылка.Организация.Наименование)
    vn_obr.Print4(selection.Ссылка, r"C:\Users\dev-miir-rk\PycharmProjects\AutoDocAccounting\rtu_upd", 1)
    vn_obr.Print4(selection.Ссылка, r"C:\Users\dev-miir-rk\PycharmProjects\AutoDocAccounting\rtu_upd", 2)

    query.Text = f"""
                             ВЫБРАТЬ
                            ООО_CMR.Ссылка КАК Ссылка
                        ИЗ
                            Документ.ООО_CMR КАК ООО_CMR
                        ГДЕ
    	                    ООО_CMR.Номер = &Номер
                          """

    query.SetParameter("Номер", "097297-29")
    selection = query.Execute().Choose()
    if selection.Next():
        print("1")
        print(selection)

    vn_obr.Print(selection.Ссылка, r"C:\Users\dev-miir-rk\PycharmProjects\AutoDocAccounting\cmr", 1)
    vn_obr.Print(selection.Ссылка, r"C:\Users\dev-miir-rk\PycharmProjects\AutoDocAccounting\cmr", 2)
