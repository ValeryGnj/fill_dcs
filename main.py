import openpyxl
from docxtpl import DocxTemplate


def main():
    # открываем файл с данными в excel
    wb = openpyxl.load_workbook('apartment_rental_agreement.xlsx', data_only=True)
    sheet = wb['Список_клиентов']

    header_list = [item.value for row in sheet['B1':'AQ1'] for item in row if
                   item.value is not None and item.value != " "]
    data_list = [item.value for row in sheet['B2':'AQ2'] for item in row if
                 item.value is not None and item.value != " "]
    doc = DocxTemplate('template_rent_by_flat.docx')
    key = header_list
    value = data_list
    a = {k: v for k, v in zip(key, value)}
    context = a
    doc.render(context)
    doc.save('result_for_rent.docx')


if __name__ == '__main__':
    main()