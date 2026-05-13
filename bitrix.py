"""Bitrix24 REST operations used by the automation workflow."""

import base64
import json

import requests

from .settings import BITRIX_WEBHOOK_URL as url, SERVER_PATH as serverPath

def send_msg_to_bot(msg):
    #bot = telebot.TeleBot("7619563934:AAFGbd6Oedpj61mp6QKCV9ls5jn5Y_vauBw")
    #list_id = ["669788718"]
    #for chat_id in list_id:
        #bot.send_message(chat_id=chat_id, text=msg)
    url = "https://corp.tknovosib.ru/rest/717/ipvi7vmargk9d92f/im.message.add.json"

    payload = json.dumps({
        "DIALOG_ID": f"chat20441",
        "MESSAGE": msg
    })
    headers = {
        'Content-Type': 'application/json'
    }
    response = requests.request("POST", url, headers=headers, data=payload)

def findDeal():
    url = "https://corp.tknovosib.ru/rest/717/ipvi7vmargk9d92f/crm.deal.list"

    payload = json.dumps({
        "filter": {
            "STAGE_ID": "C17:UC_NZTLLE"
        }
    })
    headers = {
        'Content-Type': 'application/json'
    }
    response = requests.request("POST", url, headers=headers, data=payload)

    if response.json()["result"]:
        return response.json()["result"]
    else:
        return []

def changeState(id, stageid):
    payload = json.dumps({
        "id": id,
        "fields": {
            #"STAGE_ID": "C17:UC_BBPEXS"
            "STAGE_ID": stageid
        }
    })
    headers = {
        'Content-Type': 'application/json'
    }
    response = requests.request("POST", url + 'crm.deal.update', headers=headers, data=payload)
    print("Изменение состояния сделки:", response)

def getFile(id, login, pw):
    global url

    payload = json.dumps({
        "FILTER": {
            "ID": int(id)
        },
        "select": [
            "UF_CRM_1600310015"
        ]
    })
    headers = {
        'Content-Type': 'application/json'
    }

    response_items = requests.request("POST", url + "crm.deal.list", headers=headers, data=payload)
    needtocheck = False

    for art_str in response_items.json()['result']:
        for item_num in art_str['UF_CRM_1600310015']:
            if not 'миир рк' in item_num.lower():
                continue
            else:
                needtocheck = True
                break

    if (needtocheck):

        dealID = id

        payload = json.dumps({
            "id": dealID
        })
        headers = {
            'Content-Type': 'application/json'
        }
        response = requests.request("POST", url + 'crm.deal.get', headers=headers, data=payload)

        try:
            file = response.json()['result']['UF_CRM_1651805106038'][0]
        except:
            #return None
            return "miir_no_doc"

        filepath = "./doc.pdf"
        payload = {
            "AUTH_FORM": "Y",
            "TYPE": "AUTH",
            "USER_LOGIN": login,
            "USER_PASSWORD": pw,
        }
        if file:
            response = requests.request("POST",
                                    f"https://corp.tknovosib.ru{file['downloadUrl']}".replace("auth=&", "?login=yes&"),
                                    data=payload)
        with open(filepath, 'wb') as f:
            f.write(response.content)

        return filepath
    else:
        return "ok"

def getCMRInfo(id):
    global url

    payload = json.dumps({
        "id": str(id)
    })
    headers = {
        'Content-Type': 'application/json'
    }
    response = requests.request("POST", url + 'crm.deal.get', headers=headers, data=payload)
    #print("Извлечение информации из сделки:", response)
    #print(response.json()["result"])
    targetServerPath = response.json()["result"]["UF_CRM_1683719576"]
    courierOrg = response.json()["result"]["UF_CRM_1575044812"]
    startCity = response.json()["result"]["UF_CRM_1575044651"]
    endCity = response.json()["result"]["UF_CRM_1575044638"]
    autoNum = response.json()["result"]["UF_CRM_1681272900997"]
    RTUDate = response.json()["result"]["UF_CRM_1626069863"]
    isPartial = response.json()["result"]["UF_CRM_1627529140"]
    if isPartial == '2240' or isPartial == 2240:
        isPartial = True
    else:
        isPartial = False


    targetServerPath = targetServerPath.replace('S:\\', serverPath)

    return targetServerPath, courierOrg, endCity, startCity, autoNum, RTUDate, isPartial

def addToBTX(upd, cmr, dealID):
    with open(upd, "rb") as f:
        base64upd = base64.b64encode(f.read()).decode('utf-8')
    with open(cmr, "rb") as f:
        base64cmr = base64.b64encode(f.read()).decode('utf-8')

    payload = json.dumps({
        "id": dealID
    })
    headers = {
        'Content-Type': 'application/json'
    }
    response = requests.request("POST", url + 'crm.deal.get', headers=headers, data=payload)
    print("Получение файлов:", response)

    fileArr = response.json()['result']['UF_CRM_1593667879']
    fileIDs = []
    for file in fileArr:
        fileIDs.append(file['id'])

    payloadFiles = []
    for fileID in fileIDs:
        payloadFiles.append({"id": fileID})
    payloadFiles.append([
        upd.replace('\\', '/').split('/')[-1],
        base64upd,
    ])
    payloadFiles.append([
        cmr.replace('\\', '/').split('/')[-1],
        base64cmr,
    ])
    payload = json.dumps({
        "entityTypeId": 2,
        "id": dealID,
        "fields": {
            "ufCrm_1593667879": payloadFiles
        }
    })
    response = requests.request("POST", url + 'crm.item.update', headers=headers, data=payload)
    print("Загрузка файлов:", response)

def updateFolderBTX(dealID, folderPath):
    payload = json.dumps({
        "id": dealID,
        "fields": {
            "UF_CRM_1701228716877": folderPath
        }
    })
    headers = {
        'Content-Type': 'application/json'
    }
    response = requests.request("POST", url + 'crm.deal.update', headers=headers, data=payload)
    print("Обновление путя клиента:", response)
