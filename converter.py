from xls2xlsx import XLS2XLSX
from datetime import datetime
startTime = datetime.now()

counter = 1

# код который преорбразует xls format to xlsx
# нужно менять только Full path to files где каунтер это номер файла иил закинуть все файлы в папку xfilesistinagdetoryadom

for i in range(100):
    # x2x = XLS2XLSX(f"/Users/eugene/PycharmProjects/SeaRates/Learning/DriversApplied11.1.2021/allcarrierdata ({counter}).xls")
    x2x = XLS2XLSX(f"/Users/eugene/PycharmProjects/SeaRates/Learning/xfilesistinagdetoryadom/Test{counter}.xls")
    maindata = x2x.to_xlsx(f"data/data{counter}.xlsx")

    counter += 1
    print(counter)
    print('for one iteration' + f'{datetime.now() - startTime}')

print(datetime.now() - startTime)
