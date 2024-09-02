
import openpyxl

def save_to_excel(results):
    if len(results) == 0:
        print("Нет данных для сохранения.")
        return

    workbook = openpyxl.Workbook()
    sheet = workbook.active
    sheet.title = 'Emails'

    # Заголовки
    sheet.append(['From', 'To', 'Subject', 'Date', 'Snippet'])

    for msg in results:
        sheet.append([msg.get('From'),
                       msg.get('To'),
                       msg.get('Subject'),
                       msg.get('Date'),
                       msg.get('Snippet')])

    workbook.save('./result/import_data.xlsx')
    print("Данные сохранены в 'import_data.xlsx'.")