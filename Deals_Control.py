from openpyxl import load_workbook  # Импортируем всё необходимое.
from openpyxl import Workbook

compare_set_old = set()  # Создаём изначальный список.
compare_set_new = set()  # Создаём список для сравнения.
ad_wb = load_workbook(filename='All_DEALS.xlsx')  # Загружаем текущую базу.
ad_ws = ad_wb['Лист1']
for i in ad_ws['A']:  # Добавляем базу в изначальный список.
    compare_set_old.add(i.value)

nd_wb = load_workbook(filename=r'C:\Users\Admin\Desktop\Источники данных\Аналитика 2.0\Битрикс.xlsx')  # Загружаем свежий Битрикс.
nd_ws = nd_wb['Битрикс']
for i in nd_ws['A']:  # Добавляем Битрикс в список для сравнения.
    compare_set_new.add(i.value)
compare_set_new.discard('ID')  # Убираем "ID".

for i in compare_set_new:  # Проверяем оставшиеся позиции второго списка на тип данных.
    if type(i) != int:
        i = int(i)

lost_deals = compare_set_old - compare_set_new  # Старая база - новая база = потеряшки.
while None in lost_deals:
    lost_deals.remove(None)

new_deals = compare_set_new - compare_set_old  # Новая база - старые сделки = новые сделки.
while None in new_deals:
    new_deals.remove(None)

print('Новые сделки: ', len(new_deals))  # Вывод количества новых сделок.
print()  # Разрыв строки.
print('Потерянные сделки: ', len(lost_deals))  # Вывод количества потерянных сделок.

for i in new_deals:  # Добавляем к основной базе новые сделки в колоночку.
    ad_ws.append([i, ])
ad_wb.save('All_DEALS.xlsx')  # Сохраняем базу данных.

lost_wb = Workbook()  # Создаём рабочую книгу под потеряшек.
lost_ws = lost_wb.active
for i in lost_deals:  # Добавляем каждую потеряшку в таблицу с новой строки.
    lost_ws.append([i, ])
lost_wb.save('LOST.xlsx')  # Сохраняем потеряшек.
