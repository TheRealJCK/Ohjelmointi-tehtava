#import pandas as pd
import json
from collections import ChainMap

#import excel2json

import xlrd 

# you need to install pandas and 
excel_path = r"C:\Users\Jimmy Chen\Downloads\ohjelmointitehtava\asiakas-data.xlsx"
#data = pd.read_excel(excel_path)
school_path = "C://Users/gr265676/Downloads/asiakas-data.xlsx"
home_path = "C://Users/Jimmy Chen/Downloads/ohjelmointitehtava/asiakas-data.xlsx"
wb = xlrd.open_workbook(school_path)
ws = wb.sheet_by_index(0)
adress_info = {}
info = {}

# json_str = data.to_json(orient='records')

# print(json_str[0])

# for name in json_str ['name']:
#     print (json_str ['name'] ['values'])


#with open('output.json', 'w') as f:

#https://www.codingem.com/python-__add__-method/


class Customer:

    def secondary_object(self):
        self.adress = Address()

    def __init__(self, Identifier, GivenName, FamilyName, CompanyName, Identification, IdentificationType, CustomerType, CustomerSubType, Mobile, Email, ApplicationsInUse, address):
        self.Identifier = Identifier
        self.GivenName = GivenName
        self.FamilyName = FamilyName
        self.CompanyName = CompanyName
        self.Identification = Identification
        self.IdentificationType = IdentificationType
        self.CustomerType = CustomerType
        self.CustomerSubType = CustomerSubType
        self.Mobile = Mobile
        self.Email = Email
        self.ApplicationsInUse = ApplicationsInUse

class Address:
    def __init__(self, StreetName, BuildingNumber, StairwellIdentification, Apartment, Postcode, Pobox, CityName, CountryCode):
        self.StreetName = StreetName
        self.BuildingNumber = BuildingNumber
        self.StairwellIdentification = StairwellIdentification
        self.Apartment = Apartment
        self.Postcode = Postcode
        self.Pobox = Pobox
        self.CityName = CityName    
        self.CountryCode = CountryCode

list_of_info = []

Applications_In_Use_list = []
Customer_data = Customer
adress_data = Customer_data.secondary_object




for i in range(ws.nrows):
    if i == 0:
        pass
    else:
        for  j in range(ws.ncols):

            # Jonkun syystä, jos solussa on vain numero eli ei ole ; merkkiä, niin se luukee sen solun floattina
            #  Korjasin sen if-lausekkeella.
            var = ws.cell_value(i, j)
            if not type(var) == str:
                var = int(var)
                var = str(var)

            if "ApplicationsInUse" in ws.cell_value(0,j): 

                # value = ws.cell_value(i,j)

                # if type(value) == float:
                #    var = str(var)
                if ";" in var:
                    for Applications in var.split(';'):
                        Applications_In_Use_list.append(Applications)
                else:
                    Applications_In_Use_list.append(var)
                setattr(Customer_data, ws.cell_value(0,j), Applications_In_Use_list)
            

            elif 'Address' in ws.cell_value(0,j):
                name = ws.cell_value(0,j).rsplit('.', 1)[-1]
                setattr(adress_data, name, var)
                #adress_data.update({ws.cell_value(0,j):ws.cell_value(i,j)})
                print("adresss")

            else:
                setattr(Customer_data, ws.cell_value(0,j), var)
                
                #Customer_date.set({ws.cell_value(0,j):ws.cell_value(i,j)})
                print("elseee")

            if j == ws.ncols-1:
                #address = Customer_date.secondary_object(adress_info)
                result = dict(Customer_data)
                list_of_info.append(Customer_data.__init__)
                list_of_info.append(adress_data.__init__)
                for info in list_of_info:
                    print(info, sep="\n")
                print(result, sep="\n")
                    
                #customer = dict(info)
                #end_result = customer.__add__(address)


#list_of_info.append(customer) 
            #           #print(ws.cell_value(i,j),end="\t")

# address = dict(adress_info)
# customer = dict(info) 

# for key, value in adress_info.items():
#     setattr(address, key, value)
# for key, value in info.items():
#     setattr(customer, key, value)
# customer.add_address(address)

path_school = "C://Users/gr265676/source/repos/Ohjelmointi-tehtava/output.json"
path_home = "C://Users/Jimmy Chen/source/repos/Ohjelmointi-tehtava/output.json"

with open(path_school, "a" ) as f:
                    x = json.dumps(list_of_info, indent=4)
                    f.write(x + ' \n')

                    print(x, "\n")    



    # def __add__(adress):
    #     return Address.__init__ + adress
    #     return Customer(self.kilos + address.Address)
    # def add_address(self, address):
    #     self.Address = address
        

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
# 
# 10-17
