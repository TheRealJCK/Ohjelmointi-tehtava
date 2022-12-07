import os
from os.path import join
import json
import xlrd 

file_name = input("Enter the name of the file (e.g. asiakas-data.xlsx): ")
location = input("Enter the disc where the file is located (character only): ")

def find_file(disc, look_for):
    
        print("Looking for the file...\nPlease wait...")
        for root, dirs, files in os.walk(disc):
                if look_for in files:
                    path = os.path.join(root, look_for)
                    print("File found!")
                    return path

location = location + ':\\'
if os.path.exists(location) == True:

    if  ".xlsx" in file_name:
        look_for_file = file_name
    else:
        look_for_file = file_name + ".xlsx"

    nd = find_file(location, look_for_file)

    try:
        wb = xlrd.open_workbook(nd)   
    except TypeError:
        print("\n File not found! \n Please check the file name and the disc location ")
        exit()
        
    ws = wb.sheet_by_index(0)

    list_of_data = []
    Applications_In_Use_list = []
    end_result = {}
    adress_info = {}
    other_info = {}

    for i in range(ws.nrows):

        if i == 0:
            pass
        else:
            for  j in range(ws.ncols):

                var = ws.cell_value(i, j)
                if not type(var) == str:
                    var = int(var)
                    var = str(var)
                if "ApplicationsInUse" in ws.cell_value(0,j): 
                    if ";" in var:
                        for Applications in var.split(';'):
                            Applications_In_Use_list.append(Applications)
                    else:
                        Applications_In_Use_list.append(var)
                        other_info[ws.cell_value(0,j)] = Applications_In_Use_list
                        Applications_In_Use_list = []
                elif 'Address' in ws.cell_value(0,j):
                    name = ws.cell_value(0,j).rsplit('.', 1)[-1]
                    adress_info[name] = var
                else:
                    other_info[ws.cell_value(0,j)] = var
                if j == ws.ncols-1:

                    end_result["Address"] = dict(adress_info)
                    end_result.update(other_info)

                    list_of_data.append(end_result)
                    end_result = {}
    with open('output.json', 'w', encoding='utf-8') as f:
        data_str = json.dumps(list_of_data, indent=4, ensure_ascii=False)
        f.write(data_str)

    print("\nAll done!")

else:
    print("The disc does not exist \n Please try again")

