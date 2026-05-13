"""Small parsing utilities without PDF dependencies."""

import datetime
import re


def convDate(date):
    date = date.replace('.', ' ')
    date = date.replace(',', ' ')
    day, month, year = date.split(' ')
    if len(day) == 1:
        day = "0" + day
    if len(month) == 2:
        return day + '.'+ month +'.' + year
    elif "янв" in month:
        return day + '.01.' + year
    elif "фев" in month:
        return day + '.02.' + year
    elif "мар" in month:
        return day + '.03.' + year
    elif "апр" in month:
        return day + '.04.' + year
    elif "июн" in month:
        return day + '.06.' + year
    elif "июл" in month:
        return day + '.07.' + year
    elif "авг" in month:
        return day + '.08.' + year
    elif "сен" in month:
        return day + '.09.' + year
    elif "окт" in month:
        return day + '.10.' + year
    elif "ноя" in month:
        return day + '.11.' + year
    elif "дек" in month:
        return day + '.12.' + year
    else:
        return day + '.05.' + year


def convDateTo1CFormat(date):
    import pytz

    date = datetime.datetime.strptime(date, '%d.%m.%Y')
    date = date.replace(tzinfo=pytz.UTC)
    return date


def convSum(docSum):
    if "," in docSum:
        wholeSum = re.split(r',\d\d', docSum)[0]
        sumFrac = docSum.split(',')[-1]
        wholeSum = wholeSum.replace('.', '')
        wholeSum = wholeSum.replace(' ', '')
        return wholeSum + '.' + sumFrac
    elif "." in docSum:
        return docSum.replace(' ', '')
    else:
        return docSum.replace(' ', '')
