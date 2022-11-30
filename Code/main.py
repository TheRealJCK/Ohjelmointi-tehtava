import pandas as pd
import xlrd 


wb = xlrd.open_workbook("C://Users/Jimmy Chen/Downloads/ohjelmointitehtava/asiakas-data.xlsx")
ws = wb.sheet_by_index(0)

class Customer:
    def __init__(self, Identifier, GivenName, FamilyName, CompanyName, Identification, IdentificationType, CustomerType, CustomerSubType, Mobile, Email, ApplicationsInUse):
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

        
for i in range(ws.nrows):
    if i == 0:
        pass
    else:
         = Customer(ws.cell_value(i, 0), ws.cell_value(i, 1), ws.cell_value(i, 2), ws.cell_value(i, 3), ws.cell_value(i, 4), ws.cell_value(i, 5), ws.cell_value(i, 6), ws.cell_value(i, 7), ws.cell_value(i, 8), ws.cell_value(i, 9), ws.cell_value(i, 10))
    print(ws.row_values(i))
    for  j in range(ws.ncols):
        if 0 <= j <= 9 or j == 18:
            Identifier = ws.row_values(i)[j]
        else:
            GivenName = ws.row_values(i)[j]
        print(ws.cell_value(i,j),end="\t")
    print('')


# 10-17
