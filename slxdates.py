import os
import openpyxl
from datetime import date

TODAY = date.today()

def get_files():
    CURRENT_DIR = os.path.dirname(os.path.abspath(__file__))
    # Find file with extension
    fileExt = r".xlsx"
    return [os.path.join(CURRENT_DIR, _) for _ in os.listdir(CURRENT_DIR) if _.endswith(fileExt)]

def load_xlsx_file_and_find_dates():
    files = get_files()
    result = []
    for file in files:
        dataframe = openpyxl.load_workbook(file)
        result, dat_active = extract_dates_for_each_file(result, dataframe)
        if len(result)<1:
            result.append("No dates found in the PAST")
    edit_file(dat_active, result)
    dataframe.save(file)
    dataframe.close()
    
def extract_dates_for_each_file(result, dataframe):
    dat_active = dataframe.active
    for row in range(0, dat_active.max_row):
        for col in dat_active.iter_cols(1, dat_active.max_column):
            if col[row].value == None or type(col[row].value) == str:
                continue
            if col[row].value.date()<TODAY:
                result.append(
                    f"The date from {col[row].value.strftime('%d/%m/%Y')} is in the PAST"
                )
    return result, dat_active

def edit_file(dat_active, result):    
    # Get european format
    today = TODAY.strftime('%d/%m/%Y')
    print("Today is: ", today)
    res = '. '.join([str(elem) for elem in result])
    print("RESULT---> ", res)
    # Sheet is the SheetName where the data has to be entered
    dat_active.cell(row=1, column=1).value = res

if __name__ == '__main__':
    load_xlsx_file_and_find_dates()