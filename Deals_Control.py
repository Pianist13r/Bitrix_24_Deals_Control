from openpyxl import load_workbook  # Импортируем всё необходимое.
from openpyxl import Workbook

compare_list_1 = []  # Создаём изначальный список.
compare_list_2 = []  # Создаём список для сравнения.
lost_deals = []  # Создаём список потерянных сделок.
dbwb = load_workbook(filename='All_DEALS.xlsx')  # Загружаем текущую базу.
dbws = dbwb['Лист1']
for i in dbws['A']:  # Добавляем базу в изначальный список.
    compare_list_1.append(i.value)

b_swb = load_workbook(filename=r'C:\Users\Admin\Desktop\Источники данных\Аналитика 2.0\Битрикс.xlsx')  # Загружаем
# свежий Битрикс.
b_sws = b_swb['Битрикс']
for i in b_sws['A']:  # Добавляем Битрикс в список для сравнения.
    compare_list_2.append(i.value)
compare_list_2.remove('ID')  # Убираем "ID".

for i in compare_list_1:  # Проходим по каждому элементу изначального списка.
    if i in compare_list_2:  # Если есть в новой базе - удаляем из новой (значит всё ок с ней).
        compare_list_2.remove(i)
    else:  # Если нет в новой базе - добавляем к списку потеряных.
        lost_deals.append(i)

new_deals = compare_list_2  # Всё, что осталось во втором списке, отсутствовало в первом, значит, это новые сделки.

print('Новые сделки: ', len(new_deals))  # Вывод количества новых сделок.
print()  # Разрыв строки.
print('Потерянные сделки: ', len(lost_deals) - 1)  # Вывод количества потерянных сделок(первое значение,
# по непонятной причине, - None).

for i in new_deals:  # Добавляем к основной базе новые сделки в колоночку.
    dbws.append([i, ])
dbwb.save('All_DEALS.xlsx')  # Сохраняем базу данных.

lost_wb = Workbook()  # Создаём рабочую книгу под потеряшек.
lost_ws = lost_wb.active
for i in lost_deals:  # Добавляем каждую потеряшку в таблицу с новой строки.
    lost_ws.append([i, ])
lost_wb.save('LOST.xlsx')  # Сохраняем потеряшек.
