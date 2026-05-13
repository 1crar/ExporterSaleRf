import copy
from math import *
from pdfminer.high_level import extract_pages
from pdfminer.layout import *
from .fields import MinedField as MF

coordEpsilon = 25

def roundToBase(x, base):
    '''
    Округляет число до ближайшего кратного основанию base
    '''
    return base * round(x / base)

def roundArrToBase(arr, base):
    for i in range(len(arr)):
        arr[i].x0 = roundToBase(floor(arr[i].x0), base)
        arr[i].y0 = roundToBase(floor(arr[i].y0), base)
    return arr

def sortFields(extractedFields, reverseY=False, pageSkip=3000) -> list[list[MF]]:
    '''
    Сортирует поля из документа в семантическом порядке (сверху вниз и слева направо), а также сливает
    текст, расположенный достаточно ровно, в цельные строки
    :param extractedFields: Извлеченные поля
    :param reverseY: Сортировать снизу вверх
    :param pageSkip: Пропуск между страницами. Нужен в электронных пдф, так как на каждой странице x, y сбрасываются
    :return:
    '''

    for page in extractedFields:
        page.sort(key=lambda field: (field.y0 * (-1 if reverseY else 1), field.x0))

    for pageNum, page in enumerate(extractedFields):
        for i in range(len(page)):
            page[i].y0 += pageSkip * pageNum

    mergedArr = []
    foundSingleLine = True
    fieldsCpy = copy.copy(extractedFields)

    for page in fieldsCpy:
        foundSingleLine = True
        while foundSingleLine:  # Пока можем соединить слова в строку - продолжаем повторять
            foundSingleLine = False
            i = 0
            while i < len(page) - 1:
                if abs(page[i + 1].y0 - page[i].y0) == 0:     # Слова на одной строке?
                    page[i].text = f"{page[i].text} {page[i + 1].text}" if page[i + 1].x0 >= page[i].x0 else\
                                   f"{page[i + 1].text} {page[i].text}"
                    page[i].text = page[i].text.replace("  ", " ").strip()

                    if page[i + 1].x0 >= page[i].x0:  # Новое слово идет после данного?
                        del page[i + 1]
                    else:
                        del page[i]
                        i = max(0, i - 1)
                    foundSingleLine = True
                else:
                    i += 1
        mergedArr.extend(el for el in page)
    return mergedArr

def pdfGetFields(pdf, coordEps):
    pdfPages = [page for page in extract_pages(pdf, laparams=LAParams(boxes_flow=0.25, line_margin=0.26,
                                                                      char_margin=1))]
    extractedFields = []
    for page in pdfPages:
        extractedFields.append([])

        for elem in page:
            if isinstance(elem, LTText) and len(elem.get_text().replace(" ", "").removesuffix('\n')) != 0:
                anchor = elem.y0
                if abs(elem.y0 - elem.y1) >= coordEpsilon:
                    anchor = (elem.y0 + elem.y1) / 2
                extractedFields[-1].append(MF(elem, roundToBase(floor(elem.x0), coordEps),
                                              roundToBase(floor(anchor), coordEps),
                                              elem.width, elem.height,
                                              elem.get_text()))
    return extractedFields