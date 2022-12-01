import pandas as pd
import json
from collections import ChainMap

import excel2json

import xlrd 

# you need to install pandas and 
excel_path = r"C:\Users\Jimmy Chen\Downloads\ohjelmointitehtava\asiakas-data.xlsx"
data = pd.read_excel(excel_path)

wb = xlrd.open_workbook("C://Users/Jimmy Chen/Downloads/ohjelmointitehtava/asiakas-data.xlsx")
ws = wb.sheet_by_index(0)
adress_info = {}
info = {}

# json_str = data.to_json(orient='records')

# print(json_str[0])

# for name in json_str ['name']:
#     print (json_str ['name'] ['values'])


#with open('output.json', 'w') as f:
for i in range(ws.nrows):
    if i == 0:
        pass
    else:
        for  j in range(ws.ncols):

            if 0 <= j <= 9 or j == 18:
                info.update({ws.cell_value(0,j):ws.cell_value(i,j)}) 
            else:
                adress_info.update({ws.cell_value(0,j):ws.cell_value(i,j)})

            if j == ws.ncols-1:

                end_result = [ adress_info, info ]
                #adress_info.update(info)

                with open("C://Users/Jimmy Chen/source/repos/Ohjelmointi-tehtava/output.json", "a" ) as f:
                    x = json.dumps(end_result, indent=4)
                    f.write(x + ' \n')

                print(x, "\n")


        #print(ws.cell_value(i,j),end="\t")

    

# for index in range(len(json_str)):
#     for key in json_str[index]:
#         print(json_str[index][key], sep=' ', end="\n")
#print(json_str)
#
# 
# print(data.head())

# l = data.values.tolist()

# print(l)



# dict = data.to_dict()


# for 
# xls = data.parse(data.sheet_names[0])
# print(xls.to_dict())



# foreach (var i in list)
# {
#     tupleList.Add(Tuple.Create(i.Key, i.Value));
# }






# print(dict)

# # other
# Identifier = list(data.Identifier)
# name = list(data.GivenName)
# fammily_name = list(data.FamilyName)
# company = list(data.CompanyName)
# id = list(data.Identification)
# id_type = list(data.IdentificationType)
# customer_type = list(data.CustomerType)
# customer_sub_type = list(data.CustomerSubType)
# mobile = list(data.Mobile)
# email = list(data.Email)
# application_in_use = list(data.ApplicationInUse)

# # adress
# address_street_name = list(data.Address.StreetName)
# address_building_number = list(data.Address.BuildingNumber)
# address_stairwell_id = list(data.Address.StairwellIdentification)
# address_apartment = list(data.Address.Apartment)
# address_postcode = list(data.Address.Postcode)
# address_pobox = list(data.Address.Pobox)
# address_city = list(data.Address.CityName)
# address_contry_code = list(data.Address.CountryCode)


# address_info = { }
# random_info = { }





# class Customer:
#     def __init__(self, Identifier, GivenName, FamilyName, CompanyName, Identification, IdentificationType, CustomerType, CustomerSubType, Mobile, Email, ApplicationsInUse):
#         self.Identifier = Identifier
#         self.GivenName = GivenName
#         self.FamilyName = FamilyName
#         self.CompanyName = CompanyName
#         self.Identification = Identification
#         self.IdentificationType = IdentificationType
#         self.CustomerType = CustomerType
#         self.CustomerSubType = CustomerSubType
#         self.Mobile = Mobile
#         self.Email = Email
#         self.ApplicationsInUse = ApplicationsInUse





# class Address:
#     def __init__(self, StreetName, BuildingNumber, StairwellIdentification, Apartment, Postcode, Pobox, CityName, CountryCode):
#         self.StreetName = StreetName
#         self.BuildingNumber = BuildingNumber
#         self.StairwellIdentification = StairwellIdentification
#         self.Apartment = Apartment
#         self.Postcode = Postcode
#         self.Pobox = Pobox
#         self.CityName = CityName
#         self.CountryCode = CountryCode

        
# 


# 10-17
