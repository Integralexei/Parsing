# coding=utf-8
"""
Задание: со страницы http://bcs-express.ru/kotirovki-i-grafiki выгрузить в excel список инструментов и цену из рубрик:
Голубые фишки (ММВБ), Мировые индексы, Товарные рынки
"""
import lxml.html as html
import xlsxwriter


def prices():
    main_domain_stat = 'http://bcs-express.ru/kotirovki-i-grafiki'

    page = html.parse(main_domain_stat)  # html объект для парсинга

    prices_list = []

    for j in range(6):
        e = page.getroot(). \
            find_class('cont_box cols4'). \
            pop(j)  # получаем корневой элемент нашего документа, получаем все элементы класса

        t = e.getchildren()[1]
        for elem in t:
            pr = elem.cssselect('li')
            g = [pri for pri in pr]
            for elemm in g:
                price = elemm.cssselect('p')[0]
                profit_per_day = elemm.cssselect('p')[1]
                name = elemm.cssselect('a')[0]
                prices_dict = {'instr': name.text, 'price': price.text, 'change': profit_per_day.text}
                prices_list.append(prices_dict)
    return prices_list


def export_excel(filename, prices_list):
    workbook = xlsxwriter.Workbook(filename)  # создаем новый файл Excel
    worksheet = workbook.add_worksheet()      # создаем новый рабочий листок

    bold = workbook.add_format({'bold': True})  # жирный шрифт
    field_names = ('Инструмент', 'Цена', 'Изменение')
    for i, field in enumerate(field_names):  # создали 3 поля
        worksheet.write(0, i, field, bold)  # (row,column,*args(то что мы добавляем),style)

    fields = ('instr', 'price', 'change')  # заполнение полей
    for row, price in enumerate(prices_list, start=1):
        for col, field in enumerate(fields):
            worksheet.write(row, col, price[field])  # price[field] == prices_list['instr']-извлекаем значение по ключу
    workbook.close()


def main():
    prices_e = prices()
    export_excel('pricesBKS.xlsx', prices_e)


if __name__ == '__main__':
    main()








