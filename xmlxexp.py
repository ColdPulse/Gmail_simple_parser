from openpyxl import Workbook

def save_to_excel(data):
    wb = Workbook()
    ws = wb.active
    ws.append(['From', 'Subject', 'Date'])  # Заголовки

    for item in data:
        ws.append([item['From'], item['Subject'], item['Date']])

    wb.save('./result/emails.xlsx')