"""Application entry point for the Bitrix -> PDF -> 1C automation."""

import datetime

import win32com.client

from .bitrix import (
    addToBTX,
    changeState,
    findDeal,
    getCMRInfo,
    getFile,
    send_msg_to_bot,
    updateFolderBTX,
)
from .onec.documents import operations1C
from .parsing import documents as document_parsing
from .parsing.documents import extractBitrixDocInfo


def main() -> None:
    deals = findDeal()
    #connection = win32com.client.Dispatch("V83.COMConnector").Connect('Srvr="1c-dev";Ref="logist_2025_rbc";'
    #                                                                  'Usr="РарусRPA";Pwd="7zevecoC";')
    '''catalog = getattr(connection.Catalogs, "НомераГТД")
    countries = getattr(connection.Catalogs, "СтраныМира")

    country = countries.FindByDescription("КИТАЙ")

    new_gtd = catalog.CreateItem()
    new_gtd.ТипНомераГТД = connection.Перечисления.ТипыНомеровГТД.НомерГТД
    new_gtd.РегистрационныйНомер = "10132160/010226/5040743"
    new_gtd.Код = "10132160/010226/5040743"

    print("before write:")
    print("Наименование =", repr(new_gtd.Наименование))
    print("РегистрационныйНомер =", repr(new_gtd.РегистрационныйНомер))

    new_gtd.Write()

    obj = new_gtd.Ref.GetObject()

    print("after write:")
    print("Код =", repr(obj.Код))
    print("Наименование =", repr(obj.Наименование))
    print("РегистрационныйНомер =", repr(obj.РегистрационныйНомер))'''
    #print(getattr(connection.Catalogs, "НомераГТД").FindByAttribute("РегистрационныйНомер", "10132160/010226/5040743").Код)

    #deals = None
    #testupdprint()

    if not deals:
        exit()
    else:
        for deal in deals:
            dealID = deal["ID"]
            print(dealID)
            #if not "100724" in deal["ID"]:
            #    continue
            # 95194 не извлекаются товары

            docdealfile = getFile(dealID, "papyc@tknovosib.ru", "ntE9O8PD2||C")
            #if not docdealfile:
                #print(f'По сделке {dealID} не хватает разрешительных документов. Пропускаю сделку.')
                #send_msg_to_bot(f'По сделке {dealID} не хватает разрешительных документов. Пропускаю сделку.')
                #changeState(dealID, "C17:UC_BBPEXS")
                #continue
            if docdealfile is None:
                # Нет документа в «Разрешениях» и это НЕ случай МИИР РК — пропускаем.
                print(f'По сделке {dealID} не хватает разрешительных документов. Пропускаю сделку.')
                send_msg_to_bot(f'По сделке {dealID} не хватает разрешительных документов. Пропускаю сделку.')
                changeState(dealID, "C17:UC_BBPEXS")
                continue
            if docdealfile == "miir_no_doc":
                # МИИР РК в коде ТНВЭД, но документ не вложен в «Разрешения».
                # Если сделка без перепродажи — это не ошибка (проверяется далее в 1С).
                print(f'Сделка {dealID}: МИИР РК без документа разрешения. Продолжаю — проверю перепродажу в 1С.')
                send_msg_to_bot(
                    f'Сделка {dealID}: МИИР РК без документа разрешения. Продолжаю — проверю перепродажу в 1С.')

            foundServerPath, courierOrg, endCity, startCity, autoNum, RTUDate, isPartial = getCMRInfo(dealID)
            RTUDate = datetime.datetime.strptime(RTUDate[:10], '%Y-%m-%d')


            if isPartial:
                send_msg_to_bot(
                    f'По сделке {dealID} указана частичная отгрузка. Пропускаю сделку.')
                changeState(dealID, "C17:UC_BBPEXS")
                continue
            print("Извлеченная информация:", dealID, foundServerPath, courierOrg, endCity, startCity)

            #connection_1c = win32com.client.Dispatch("V83.COMConnector").Connect('Srvr="1c-dev";Ref="logist_2025_rbc";'
            #                                                                      'Usr="РарусRPA";Pwd="7zevecoC";')
            connection_1c = win32com.client.Dispatch("V83.COMConnector").Connect('Srvr="sql-1c";Ref="logist_2025";'          
                                                                                  'Usr="РарусRPA";Pwd="7zevecoC";')
            clientServerPath = None
            try:
                #if docdealfile:
                if docdealfile and docdealfile != "miir_no_doc":
                    try:
                        orgName, items, itemCodes = extractBitrixDocInfo('./doc.pdf')
                    except:
                        print('Скип разрешения')
                        send_msg_to_bot(f"Предупреждение: Не могу проеверить разрешения для сделки {dealID}. Возможно, разрешения не МИИР РК. Пробую обработать сделку")
                    #if not check1Cdocdeal(dealID, orgName, itemCodes, connection_1c):
                        #send_msg_to_bot(f'По сделке {dealID} не сходятся данные в разрешительных документах и в заказе клиента. Проверьте, пожалуйста.')

                document_parsing.dealID = dealID
                updNum, updDate, updSumItems, updSumNDS, updSum, gtdArr, dfUpd = document_parsing.extractServerDocInfo(foundServerPath)
                print(gtdArr)

                if updNum is None:
                    send_msg_to_bot(
                        f'УПД сделки {dealID} либо не найдено, либо имеет недопустимый формат (скан)')
                    print(f"УПД сделки {dealID} либо не найдено, либо имеет недопустимый формат (скан)\n")
                    changeState(dealID, "C17:UC_BBPEXS")
                    continue
                if updNum == "" or updDate == "" or updSumItems == "" or updSum == "":
                    send_msg_to_bot(
                        f'УПД сделки {dealID} либо не найдено, либо имеет недопустимый формат (скан)')
                    changeState(dealID, "C17:UC_BBPEXS")
                    print(f"Не удалось извлечь полную информацию из УПД сделки {dealID} (неверная кодировка текста или структура документа)\n")
                    continue
                changeState(dealID, "C17:UC_BBPEXS")
                clientServerPath = operations1C(dealID, updSum, updNum, updDate, connection_1c, courierOrg, endCity, startCity, gtdArr, dfUpd, RTUDate, autoNum)
            except Exception as e:
                send_msg_to_bot(
                    f'Со сделкой {dealID} неизвестные проблемы, возьмите в ручную работу, пожалуйста')
                print(f"Со сделкой {dealID} неизвестные проблемы, возьмите в ручную работу, пожалуйста")
                print(e)
                #changeState(dealID, "C17:UC_BBPEXS")
                continue
            if not clientServerPath:
                print(f"Документы не были созданы из-за внутренней ошибки при выполнении операций с 1С\n")
                continue
            print(clientServerPath)
            updateFolderBTX(dealID, clientServerPath)
            try:
                addToBTX(clientServerPath + '/rtu_upd.pdf', clientServerPath + '/cmr.pdf', dealID)
            except:
                print('Не получилось что-то с УПД')
            changeState(dealID, "C17:UC_DHCK5O")

