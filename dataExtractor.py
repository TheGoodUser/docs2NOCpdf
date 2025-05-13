from pandas import read_excel
import pdfGenerator as pdfg # mine generated one
from chooseFile import choose_file
from tkinter import messagebox as msgbx

try:
    dataframe = read_excel(f'{choose_file()}')
except:
    msgbx.showinfo("Excel2PDF", "Select a file to generate NOC's!")
    exit()


# for f in range(dataframe.columns)
def dataRetrieverFromExcel():

    for f in range(0, dataframe.shape[0]):
        # date formatting in DD/MM/YYYY
        '''date = f'{dataframe.iloc[f, 10]}'[:-9].split("-")
        rep = date[0]
        date[0] = date[2]
        date[2] = rep
        date = "/".join(date)'''
        # Safely access and format the date from the dataframe
        date_string = f'{dataframe.iloc[f, 10]}'[:-9]
        date_parts = date_string.split("-")

        if len(date_parts) == 3:  # Ensure date has exactly 3 parts
            # Swap day and year to get dd/mm/yyyy
            date_parts[0], date_parts[2] = date_parts[2], date_parts[0]
            date = "/".join(date_parts)
        else:
            date = "Invalid date"  # Handle invalid date format

        print(date)



        # vehicle no. format CG-22-CF-2222 or CG-22-F-2222
        vehicleDetail = f'{dataframe.iloc[f, 3]}'
        _0 = vehicleDetail[:2]
        _1 = vehicleDetail[2:4]
        _3 = vehicleDetail[4:6]
        _3 = "".join([strOnly for strOnly in [*_3] if strOnly.isalpha()])
        _4 = vehicleDetail[-4:]
        i = vehicleDetail
        vehicleDetail = f"{_0}-{_1}-{_3}-{_4}"


        data = [dataframe.iloc[f, 1], dataframe.iloc[f, 2], vehicleDetail, dataframe.iloc[f, 5], dataframe.iloc[f, 6], date]
        # print(data)
        pdfg.generatePDF(dataframe.iloc[f, 1], dataframe.iloc[f, 2], vehicleDetail, dataframe.iloc[f, 5], dataframe.iloc[f, 6], date)



