import openpyxl
from docxtpl import DocxTemplate


def main():
    # открываем файл с данными в excel
    wb = openpyxl.load_workbook('apartment_rental_agreement.xlsx', data_only=True)
    # получаем лист с данными
    sheet = wb['Список_клиентов']
    # генерируем список заголовков столбцов с данными
    header_list = [item.value for row in sheet['B1':'AQ1'] for item in row if
                   item.value is not None and item.value != " "]
    # генерируем список данных из строки таблицы
    data_list = [item.value for row in sheet['B2':'AQ2'] for item in row if
                 item.value is not None and item.value != " "]
    # открываем шаблон в word для заполнения данными из excel
    doc = DocxTemplate('template_rent_by_flat.docx')

    # присваиваем ключам и значениям соответствующие списки с данными
    key = header_list
    value = data_list
    # генерируем словарь по ключам и значениям из списков сгенерированных выше
    a = {k: v for k, v in zip(key, value)}

    # заполняем шаблон данными из словаря
    context = a
    doc.render(context)
    # сохраняем готовый документ
    doc.save('result_for_rent.docx')


if __name__ == '__main__':
    main()