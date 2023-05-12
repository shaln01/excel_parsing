import pandas as  pd
import numpy as np
import copy
import openpyxl
import json
df = pd.read_excel("C:\\Users\\ShailendraSingh\\Desktop\\81472_11282022_1220146211.xlsx",sheet_name="FHLMC LLPAs")
df_sec_col = df.iloc[:,1]
print (type(df_sec_col))
all_data_list = []
change_fico_data = False
main_col_head = ["FICO","Product Feature"]
col_fico_dict = {"fico_idx": [], "fico_cols": [[]]}
print("hello",df_sec_col)
counter = 0
counter = 0
for row in df_sec_col:
    if row in main_col_head:
        if len(col_fico_dict["fico_idx"])>=1:
            col_fico_dict["fico_idx"].append(counter)
            all_data_list.append(copy.deepcopy(col_fico_dict))
            col_fico_dict = {"fico_idx": [], "fico_cols": [[]]}
        change_fico_data = True
        col_fico_dict["fico_idx"].append(counter)
        col_fico_dict["fico_cols"][0].append(row)
    if row not in main_col_head and row is not np.nan and change_fico_data:
        col_fico_dict["fico_cols"][0].append(row)
    if change_fico_data and row is np.nan:
        col_fico_dict["fico_idx"].append(counter)
        all_data_list.append(copy.deepcopy(col_fico_dict))
        col_fico_dict = {"fico_idx": [], "fico_cols": [[]]}
        change_fico_data = False
    counter += 1
    if len(all_data_list)>=3:
        break

ltv_cols = df.iloc[:,2:]

for col in ltv_cols:
    for idx_range in all_data_list:
        col_data = list(df[col][idx_range["fico_idx"][0]:idx_range["fico_idx"][1]])
        idx_range["fico_cols"].append(col_data)

table_names = ["Fico15", "Cash Out", "Product Features"]
json_list = [{}]
for outer_idx, fico_data in enumerate(all_data_list):
    json_out_dict = {
        "FICO": [],
        "LTV": []
    }
    for idx, data in enumerate(fico_data["fico_cols"]):
        if idx==0:
            json_out_dict["FICO"] = data[1:]
            continue
        ltv_values = {
            "min": "",
            "max": "",
            "values": []
        }
        ltv_head = data[0]
        if "-" in ltv_head:
            ltv_min,ltv_max = ltv_head.split("-")
        else:
            ltv_min = 0
            ltv_max = ""
            for letter in ltv_head:
                if letter.isnumeric():
                    ltv_max += letter

        ltv_values["min"] = ltv_min
        ltv_values["max"] = ltv_max
        ltv_val_df = pd.DataFrame(data[1:])
        print(ltv_val_df.fillna("0",inplace=True))
        print(ltv_val_df[0].to_list())
        ltv_values["values"] = ltv_val_df[0].to_list()
        json_out_dict["LTV"].append(copy.deepcopy(ltv_values))
        if idx==len(fico_data["fico_cols"])-1:
            json_list[0][table_names[outer_idx]] = copy.deepcopy(json_out_dict)

json_object = json.dumps(json_list, indent=4)
print(json_object)


with open("json_data.json","w") as j_file:
    print("Hello")
    j_file.write(json_object)
# print(json_list)
