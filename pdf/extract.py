import camelot
import pdfminer
import PyPDF2
from PyPDF2 import PdfReader, PdfWriter

from .recognize import pdfGetFields, sortFields


def get_info(pdf_path, rotated=False, eps = 7):
    '''
    Извлекает весь текст и все возможные таблицы из документа.
    Текст соединяется строками, если слова расположены достаточно ровно (т.е. мало варьируют положение по Y)
    Иногда фрагменты текста расположены в вертикальном порядке, т.е. как заголовки таблицы. Из-за этого простой поиск
    не работает при извлечении целевой информации. Для этого текст извлекается еще в форме таблиц, где это возможно
    :param img_path: Путь к pdf
    :return:
    '''
    minEpsilon = eps
    # imgH, imgW = pdf2image.convert_from_path(pdf_path)[-1].size
    reader = PyPDF2.PdfReader(pdf_path)
    imgH, imgW = reader.pages[0].mediabox.height, reader.pages[0].mediabox.width,
    hEps = max(minEpsilon, imgH // 100)
    allTables2D = []

    if rotated:
        reader = PdfReader(pdf_path)
        writer = PdfWriter()
        for page in reader.pages:
            page.rotate(90)
            writer.add_page(page)
        pdf_path = "./rotated.pdf"
        with open(pdf_path, "wb") as pdf_out:
            writer.write(pdf_out)

    pagesArr = pdfGetFields(pdf_path, hEps)
    for page in pagesArr:
        for field in page:
            field.text = field.text.replace("\n", "")
    fieldsArr = sortFields(pagesArr, reverseY=True, pageSkip=imgH)
    tables = camelot.read_pdf(pdf_path, line_scale=100)
    for i in range(len(tables)):
        allTables2D.append(tables[i].df.values.tolist())
    return fieldsArr, allTables2D