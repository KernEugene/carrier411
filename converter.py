from xls2xlsx import XLS2XLSX

counter = 1

for i in range(50):
    x2x = XLS2XLSX(f"/Users/eugene/PycharmProjects/SeaRates/Learning/attachments/allcarrierdata ({counter}).xls")
    x2x.to_xlsx(f"data{counter}.xlsx")
    counter += 1