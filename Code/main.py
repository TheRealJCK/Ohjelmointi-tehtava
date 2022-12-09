import json
import pandas as pd


test_path = r"C:\Users\Jimmy Chen\Downloads\anime (version 1).xlsb.xlsx"
file_name = input("Enter the full file path (e.g. C:/Users\Jimmy Chen\Downloads\asiakas-data.xlsx): ")

# read the file path from the input
try:
    df = pd.read_excel(file_name)
except FileNotFoundError:
    print("\n File not found! \n Please check the file path ")
    exit()

# replace the NaN values with empty strings
df.fillna('', inplace=True)
result = df.to_dict(orient='records')

print(result)
print("\nConverting to JSON...")

list_of_data = []
end_result = {}
Applications_In_Use_list = []
adress_info = {}
other_info = {}

for i in range(len(result)):

    for key, value in result[i].items():

        # conver the value to string if it is not a string
        if not type(value) == str:
            value = int(value)
            value = str(value)

        # if the key is address, then make a dictionary of the address info
        if "Address" in key:
            key = key.rsplit('.', 1)[-1]
            adress_info[key] = value

        # if the key is ApplicationsInUse, then make a list of the applications in use
        elif "ApplicationsInUse" in key:

            if ";" in value:
                for Applications in value.split(';'):
                    Applications_In_Use_list.append(Applications)
            else:
                Applications_In_Use_list.append(value)
            other_info[key] = Applications_In_Use_list
            Applications_In_Use_list = []
        
        # if the key is not address or ApplicationsInUse, then add it to the other info dictionary
        else:
            other_info[key] = value

        if len(other_info) == 11:
            end_result["Address"] = dict(adress_info) 
            end_result.update(other_info)
            list_of_data.append(end_result)
            end_result = {}
            other_info = {}

with open('output.json', 'w', encoding='utf-8') as f:
    data_str = json.dumps(list_of_data, indent=4, ensure_ascii=False)
    f.write(data_str)


print("\nAll done!")
