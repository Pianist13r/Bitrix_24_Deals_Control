# Bitrix_24_Deals_Control
Короче чтоб смотреть по сделкам в битриксе в твоей выборке сколько появилось новых, а сколько пропало (если кто-то п****т сделки). По пропавшим автоматом генерируется файл.
Суть в чём:
Изначально в коде есть одна единственная ссылка на файл Битрикс.xlsx (которую легко поменять) в котором лист должен называться "Битрикс".
В этом файле преполагается, что в колонке A содержаться ID сделок. Это тот отчёт, который вы выгружаете регулярно и как раз проверяете.
При первом включении программы база заполняется автоматически, дозаполняется также. Хранится в файле All_DEALS.xlsx в папке с программой.
После запуска в консоли программа, слегка подумав, в зависимости от размеров ваших файлов, выдаст:
1) Количество новых сделок, появившихся со времени последней проверки.
2) Количество пропавших сделок. Пропавший означает, что в базе числится этот ID, а в текущем скачаном .xls его нет.
3) Перезаписывает файл LOST.xlsx, в котором помещает все потерянные ID в столбик.
Если вы хотите удалить какую-то сделку из базы, чтобы на его отсутствие программа перестала срабатывать, просто удалите строчку из файла All_DEALS.xlsx с этим ID.
