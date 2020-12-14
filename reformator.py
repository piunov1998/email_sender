import win32com.client

Excel = win32com.client.Dispatch("Excel.Application")
data = Excel.Workbooks.Open(u'C:\\Users//Тоха//github//email_sender//userdata//data.xlsx')
sheet = data.ActiveSheet

emails = []
i = 0
go = True

while go:
    i += 1
    cell = sheet.Cells(i, 1).value
    if sheet.Cells(i, 1).value == None:
        go = False
    if '@' in str(cell):
        emails.append(cell)
data.Close()

data = data = Excel.Workbooks.Open(u'C:\\Users//Тоха//github//email_sender//userdata//data_new.xlsx')
sheet = data.ActiveSheet

i = 0
for email in emails:
    i += 1
    sheet.Cells(i, 1).value = email
data.Save()
data.Close()
Excel.Quit()