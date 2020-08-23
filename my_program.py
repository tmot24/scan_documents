import os
import re
import time

import cv2
import pandas as pd
import pytesseract


class Invoice:
    def __init__(self, document, number=None, date=None, tax_free=None, tax=None, total=None,
                 consignee_name=None, consignee_address=None, nomenclature=None):
        self.consignee_address = consignee_address
        self.document = document
        self.number = number
        self.date = date
        self.tax_free = tax_free
        self.tax = tax
        self.total = total
        self.consignee_name = consignee_name
        self.nomenclature = nomenclature


class Waybill:
    def __init__(self, document, number=None, date=None, volume=None, tax_free=None, tax=None, total=None):
        self.total = total
        self.tax = tax
        self.tax_free = tax_free
        self.volume = volume
        self.document = document
        self.number = number
        self.date = date


def func_invoice_read(document):
    regex_continue = r"(продолжение)"
    matches_continue = re.search(pattern=regex_continue, string=document, flags=re.IGNORECASE)

    if matches_continue:
        regex_sum = r"(\w+\s+к\s+оплате:\D+)(\S+)(\D+)(\S+)(\D+)(\S+)"
        matches_sum = re.search(pattern=regex_sum, string=document, flags=re.IGNORECASE)
        if matches_sum:
            list_invoice_tax_free.pop()
            list_invoice_tax.pop()
            list_invoice_total.pop()
            invoice.tax_free = matches_sum.group(2)
            invoice.tax = matches_sum.group(4)
            invoice.total = matches_sum.group(6)
            list_invoice_tax_free.append(invoice.tax_free)
            list_invoice_tax.append(invoice.tax)
            list_invoice_total.append(invoice.total)

    else:
        regex_sum = r"(\w+\s+к\s+оплате:\D+)(\S+)(\D+)(\S+)(\D+)(\S+)"
        matches_sum = re.search(pattern=regex_sum, string=document, flags=re.IGNORECASE)
        if matches_sum:
            invoice.tax_free = matches_sum.group(2)
            invoice.tax = matches_sum.group(4)
            invoice.total = matches_sum.group(6)
        else:
            invoice.tax_free = file_name
            invoice.tax = file_name
            invoice.total = file_name

        list_invoice_tax_free.append(invoice.tax_free)
        list_invoice_tax.append(invoice.tax)
        list_invoice_total.append(invoice.total)

        regex_invoice = r"(СЧЕТ-ФАКТУРА)(\s+№\s+)(\d+)(\s+от\s+)(\d\d\.\d\d\.\d\d\d\d)"
        matches_invoice = re.search(pattern=regex_invoice, string=document, flags=re.IGNORECASE)
        if matches_invoice:
            invoice.number = matches_invoice.group(3)
            invoice.date = matches_invoice.group(5)
        else:
            invoice.number = file_name
            invoice.date = file_name

        list_invoice_number.append(invoice.number)
        list_invoice_date.append(invoice.date)

        regex_consignee_name = r"(.рузополучатель\s+и\s+его\s+адрес:\s+)(.+)(\".+\")(.+)(К\s+платежн)"
        matches_consignee = re.search(pattern=regex_consignee_name, string=document, flags=(re.IGNORECASE | re.DOTALL))
        if matches_consignee:
            invoice.consignee_name = matches_consignee.group(3)
            invoice.consignee_address = matches_consignee.group(4)
        else:
            invoice.consignee_name = file_name
            invoice.consignee_address = file_name

        list_invoice_consignee_name.append(invoice.consignee_name)
        list_invoice_consignee_address.append(invoice.consignee_address)

        regex_nomenclature = r"^(.+-)(.+168)"
        matches_nomenclature = re.search(pattern=regex_nomenclature, string=document, flags=re.MULTILINE)
        if matches_nomenclature:
            invoice.nomenclature = matches_nomenclature.group(1)
        else:
            invoice.nomenclature = file_name

        list_invoice_nomenclature.append(invoice.nomenclature)

    global torg_count
    torg_count = 0


def func_waybill_read(document):
    regex_waybill = r"(ТОВАРНАЯ\sНАКЛАДНАЯ)\s+(\d+)\D+(\d\d.\d\d.\d\d\d\d)"
    matches_waybill = re.search(pattern=regex_waybill, string=document, flags=re.IGNORECASE)
    if matches_waybill:
        waybill.number = matches_waybill.group(2)
        waybill.date = matches_waybill.group(3)
    else:
        waybill.number = file_name
        waybill.date = file_name

    regex_total = r"(Всего по.+\D)\D+(\d+\,\d+)\D+(\d.+,\d\d)\D+(\d.+,\d\d)\D+(\d.+,\d\d)\s"
    matches_total = re.search(pattern=regex_total, string=document, flags=re.IGNORECASE)
    if matches_total:
        waybill.volume = matches_total.group(2)
        waybill.tax_free = matches_total.group(3)
        waybill.tax = matches_total.group(4)
        waybill.total = matches_total.group(5)
    else:
        waybill.volume = file_name
        waybill.tax_free = file_name
        waybill.tax = file_name
        waybill.total = file_name

    global torg_count

    if torg_count == 0:

        reserve_list_wb_number.clear()
        reserve_list_wb_date.clear()
        reserve_list_wb_volume.clear()
        reserve_list_wb_tax_free.clear()
        reserve_list_wb_tax.clear()
        reserve_list_wb_total.clear()

        list_waybill_number.append(waybill.number)
        list_waybill_date.append(waybill.date)
        list_waybill_volume.append(waybill.volume)
        list_waybill_tax_free.append(waybill.tax_free)
        list_waybill_tax.append(waybill.tax)
        list_waybill_total.append(waybill.total)

    elif torg_count == 1:

        wb_number = list_waybill_number.pop()
        wb_date = list_waybill_date.pop()
        wb_volume = list_waybill_volume.pop()
        wb_tax_free = list_waybill_tax_free.pop()
        wb_tax = list_waybill_tax.pop()
        wb_total = list_waybill_total.pop()

        reserve_list_wb_number.append(wb_number)
        reserve_list_wb_date.append(wb_date)
        reserve_list_wb_volume.append(wb_volume)
        reserve_list_wb_tax_free.append(wb_tax_free)
        reserve_list_wb_tax.append(wb_tax)
        reserve_list_wb_total.append(wb_total)

        reserve_list_wb_number.append(waybill.number)
        reserve_list_wb_date.append(waybill.date)
        reserve_list_wb_volume.append(waybill.volume)
        reserve_list_wb_tax_free.append(waybill.tax_free)
        reserve_list_wb_tax.append(waybill.tax)
        reserve_list_wb_total.append(waybill.total)

        list_waybill_number.append(reserve_list_wb_number.copy())
        list_waybill_date.append(reserve_list_wb_date.copy())
        list_waybill_volume.append(reserve_list_wb_volume.copy())
        list_waybill_tax_free.append(reserve_list_wb_tax_free.copy())
        list_waybill_tax.append(reserve_list_wb_tax.copy())
        list_waybill_total.append(reserve_list_wb_total.copy())

    else:
        list_waybill_number.pop()
        list_waybill_date.pop()
        list_waybill_volume.pop()
        list_waybill_tax_free.pop()
        list_waybill_tax.pop()
        list_waybill_total.pop()

        reserve_list_wb_number.append(waybill.number)
        reserve_list_wb_date.append(waybill.date)
        reserve_list_wb_volume.append(waybill.volume)
        reserve_list_wb_tax_free.append(waybill.tax_free)
        reserve_list_wb_tax.append(waybill.tax)
        reserve_list_wb_total.append(waybill.total)

        list_waybill_number.append(reserve_list_wb_number.copy())
        list_waybill_date.append(reserve_list_wb_date.copy())
        list_waybill_volume.append(reserve_list_wb_volume.copy())
        list_waybill_tax_free.append(reserve_list_wb_tax_free.copy())
        list_waybill_tax.append(reserve_list_wb_tax.copy())
        list_waybill_total.append(reserve_list_wb_total.copy())

    torg_count += 1


start = time.time()
files = os.listdir('./try/')
torg_count = 0
invoice_count = 0

reserve_list_wb_number = []
reserve_list_wb_date = []
reserve_list_wb_volume = []
reserve_list_wb_tax_free = []
reserve_list_wb_tax = []
reserve_list_wb_total = []

list_invoice_date = []
list_invoice_number = []
list_invoice_tax_free = []
list_invoice_tax = []
list_invoice_total = []
list_invoice_consignee_name = []
list_invoice_consignee_address = []
list_invoice_nomenclature = []

list_waybill_date = []
list_waybill_number = []
list_waybill_volume = []
list_waybill_tax_free = []
list_waybill_tax = []
list_waybill_total = []

dict_data = {'СФ Дата': list_invoice_date,
             'СФ Номер': list_invoice_number,
             'СФ Сумма без НДС': list_invoice_tax_free,
             'СФ НДС': list_invoice_tax,
             'СФ Всего': list_invoice_total,
             'СФ Название грузополучателя': list_invoice_consignee_name,
             'СФ Адрес грузополучателя': list_invoice_consignee_address,
             'СФ Номенклатура': list_invoice_nomenclature,
             'Т12 Дата': list_waybill_date,
             'Т12 Номер': list_waybill_number,
             'Т12 Вес': list_waybill_volume,
             'Т12 Сумма без НДС': list_waybill_tax_free,
             'Т12 НДС': list_waybill_tax,
             'Т12 Всего': list_waybill_total}

for file_name in files:
    print(file_name)
    image = cv2.imread(filename=f'./try/{file_name}')

    gray_image = cv2.cvtColor(src=image, code=cv2.COLOR_BGR2GRAY)
    gaussian_image = cv2.GaussianBlur(src=gray_image, ksize=(3, 3), sigmaX=3)

    pytesseract.pytesseract.tesseract_cmd = r'D:\Program Files\Tesseract-OCR\tesseract.exe'
    custom_config = r'--psm 6'
    text = pytesseract.image_to_string(image=gaussian_image, lang='rus', config=custom_config)

    regex_invoices = r"(СЧЕТ\W+ФАКТУРА)"
    matches_invoices = re.search(pattern=regex_invoices, string=text, flags=re.IGNORECASE)
    regex_waybills = r"(ТОРГ-12)"
    matches_waybills = re.search(pattern=regex_waybills, string=text)

    if matches_invoices:
        invoice = Invoice(document=text)
        func_invoice_read(text)
        # print(text)
    elif matches_waybills:
        waybill = Waybill(document=text)
        func_waybill_read(text)
        # print(text)

print(dict_data)
invoice_df = pd.DataFrame(data=dict_data)
invoice_df.to_excel(excel_writer='data.xlsx', index=False)

end = time.time()
print(f'Обработка завершена за: {(end - start) / 60:.3f} минут')
