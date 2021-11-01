from openpyxl import load_workbook
import csv
from datetime import datetime


def write_csv(data):
    with open('cms.csv', 'a') as f:
        writer = csv.writer(f)
        writer.writerow(('LEGALNAME', data['LegalName'],
                         'EMAIL', data['Email'],
                         'CONTACT_NAME ', data['ContactName']))


def refactoring(field):
    ans = ''
    for i in str(field).split():
        ans += str(i[0].upper() + i[1::].lower()) + ' '
        continue
    return ans



def search():
    startTime = datetime.now()
    counter = 1
    for i in range(50):
        workbook = load_workbook(filename=f"data/data{counter}.xlsx")
        counter += 1
        sheet = workbook.active
        amount = sheet.max_row
        list_mails = ['@gmail.com', '@yahoo.com', '@yahoo.com.mx', '@hotmail.com', '@icloud.com', '@outlook.com',
                      '@live.com', '@aol.com', '@usa.com', '@ymail.com']


        for i in range(1, amount):
             email = (sheet.cell(row=i, column=12).value)
             legalname = (sheet.cell(row=i, column=3).value)
             contact_name = (sheet.cell(row=i, column=13).value)
             print(email)

             if email == None:
                 pass
             else:
                 data = {'LegalName': refactoring(legalname),
                          'Email': email,
                          'ContactName': refactoring(contact_name)}

                 write_csv(data)




             # for domain in list_mails:
             #     if email == None:
             #         pass
             #     elif domain in email:
             #         data = {'LegalName': refactoring(legalname),
             #                 'Email': email,
             #                 'ContactName': refactoring(contact_name)}
             #         write_csv(data)
             #     else:
             #         pass
             #     continue
    print(datetime.now() - startTime)





search()



