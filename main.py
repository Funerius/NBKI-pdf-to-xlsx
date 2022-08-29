from pdfminer.high_level import extract_text        # Импорт библиотеки, отвечающей за извлечение текста из pdf
import xlsxwriter                                   # Импорт библиотеки, отвечающей за преобразование в xlsx формат
import re                                           # Встроенная библиотека. Отвечает за регулярные выражения
import tkinter as tk                                # Встроенная библиотека. Отвечает за GUI
from tkinter import filedialog
from tkinter import messagebox

# Создание функции => Нажатие кнопки


def get_file_path():
    filepath = filedialog.askopenfilename(title="Выбрать файл",
                                          filetypes=(("pdf files", "*.pdf"), ("all files", "*.*")))

    # Достаем текст из pdf
    text = extract_text(f'{filepath}')

    # Регулярные выражения для парсинга нужной информации

    name_client = re.search(r'([А-ЯЁ-]+\s[А-ЯЁ-]+\s[А-ЯЁ-]+)', text)
    vid = re.findall(r'(Вид:.*)', text)
    limit = re.findall(r'(Размер/лимит:.*)', text)
    psk = re.findall(r'(ПСК%%:.*)', text)
    opened = re.findall(r'(Открыт:.*)', text)
    status = re.findall(r'(Статус:.*)', text)
    final_payment = re.findall(r'(Финальн.платеж:.*)', text)
    debt = re.findall(r'(Задолж-сть:.*)', text)
    next_payment = re.findall(r'(След.платеж:.*)', text)
    delay_low = re.findall(r'Просрочек от 30 до 59 дн.:.*', text)
    delay_mid = re.findall(r'Просрочек от 60 до 89 дн.:.*', text)
    delay_high = re.findall(r'Просрочек более, чем на 90 дн.:.*', text)

    # Собираем массив из полученных данных
    result = list(
        zip(vid, limit, psk, opened, status, final_payment, debt, next_payment, delay_low, delay_mid, delay_high))

    # Создание пути в ту же директорию для выгрузки xlsx

    new_filepath = filepath.split('.')[0]

    # Создание xlsx файла
    workbook = xlsxwriter.Workbook(f'{new_filepath}.xlsx', {'strings_to_numbers': True})
    worksheet = workbook.add_worksheet()

    # Форматирование стилей таблицы
    header_format = workbook.add_format({'bg_color': '#000000', 'font_color': '#FFFFFF', 'bold': True})
    cell_format = workbook.add_format({'border': True, 'bg_color': '#DCDCDC', 'align': 'left'})
    formulas_format = workbook.add_format({'bg_color': '#F5B7B1'})

    # Заголовок
    worksheet.write("B1", "ФИО", header_format)
    worksheet.merge_range("C1:L1", name_client.group(0), header_format)
    worksheet.write("B2", "ВИД", header_format)
    worksheet.set_column('B:B', 30)
    worksheet.write("C2", "РАЗМЕР/ЛИМИТ", header_format)
    worksheet.set_column('C:C', 20)
    worksheet.write("D2", "ПСК", header_format)
    worksheet.set_column('D:D', 20)
    worksheet.write("E2", "ОТКРЫТ", header_format)
    worksheet.set_column('E:E', 20)
    worksheet.write("F2", "СТАТУС", header_format)
    worksheet.set_column('F:F', 30)
    worksheet.write("G2", "ФИНАЛЬН.ПЛАТЕЖ", header_format)
    worksheet.set_column('G:G', 20)
    worksheet.write("H2", "ЗАДОЛЖ-СТЬ", header_format)
    worksheet.set_column('H:H', 20)
    worksheet.write("I2", "СЛЕД.ПЛАТЕЖ", header_format)
    worksheet.set_column('I:I', 20)
    worksheet.write("J2", "Пр. 30-59 дн", header_format)
    worksheet.set_column('J:J', 15)
    worksheet.write("K2", "Пр. 60-89 дн", header_format)
    worksheet.set_column('K:K', 15)
    worksheet.write("L2", "Пр. > 90 дн", header_format)
    worksheet.set_column('L:L', 15)

    # Основная таблица
    row = 2
    col = 1

    for vid, limit, psk, opened, status, final_payment, debt, next_payment, delay_low, delay_mid, delay_high in result:
        next_payment = next_payment.split(": ")[1]
        next_payment = next_payment.replace("RUB ", "")
        limit = limit.split(": ")[1]
        limit = limit.replace("RUB ", "")
        debt = debt.split(": ")[1]
        debt = debt.replace("RUB ", "")
        status = status.replace("Счет закрыт - переведен на ",
                                "Счет закрыт - переведен на обслуживание в другую организацию")

        worksheet.write(row, col, vid.split(": ")[1], cell_format)
        worksheet.write(row, col + 1, limit, cell_format)
        worksheet.write(row, col + 2, psk.split(": ")[1], cell_format)
        worksheet.write(row, col + 3, opened.split(": ")[1], cell_format)
        worksheet.write(row, col + 4, status.split(": ")[1], cell_format)
        worksheet.write(row, col + 5, final_payment.split(": ")[1], cell_format)
        worksheet.write(row, col + 6, debt, cell_format)
        worksheet.write(row, col + 7, next_payment, cell_format)
        worksheet.write(row, col + 8, delay_low.split(": ")[1], cell_format)
        worksheet.write(row, col + 9, delay_mid.split(": ")[1], cell_format)
        worksheet.write(row, col + 10, delay_high.split(": ")[1], cell_format)
        row += 1

    # Формулы с автосуммой
    worksheet.set_column('M:M', 20)
    worksheet.write(row + 1, col + 10, "ИТОГО: ", formulas_format)
    worksheet.write(row + 2, col + 10, "РАЗМ/ЛИМИТ: ", formulas_format)
    worksheet.write_formula(row + 3, col + 10, '=SUM(C3:C1024)', formulas_format)
    worksheet.write(row + 4, col + 10, "ЗАДОЛЖ-ТЬ: ", formulas_format)
    worksheet.write_formula(row + 5, col + 10, '=SUM(H3:H1024)', formulas_format)
    worksheet.write(row + 6, col + 10, "СЛЕД. ПЛАТЕЖ: ", formulas_format)
    worksheet.write_formula(row + 7, col + 10, '=SUM(I3:I1024)', formulas_format)

    # Закрытие xlsx файла
    workbook.close()

    # Уведомление о том, что все прошло корректно
    tk.messagebox.showinfo(title=None, message="Файл преобразован успешно")


# Основное графическое отображение программы

root = tk.Tk()
root.geometry('600x100')
root.resizable(False, False)
root.title('NBKI to xlsx')
root.iconbitmap('nbki.ico')


# Виджеты
btn_choose_file = tk.Button(root, text='Выбрать файл', command=get_file_path, width=18, height=2, bg="#DCDCDC")
lbl_filepath = tk.Label(root, text='Результат обработки будет\n '
                                   'сохранен в ту же директорию', width=40, height=3, font='Consolas 14')

# Группировка виджетов в окне

btn_choose_file.grid(row=0, column=1, padx=5, pady=10)
lbl_filepath.grid(row=0, column=0, padx=10, pady=10)

# Закрытие программы
root.mainloop()
