import pandas as  pd
import numpy as np
import copy
import sys
import openpyxl
import json
df = pd.read_excel("C:\\Users\\ShailendraSingh\\Desktop\\81472_11282022_1220146211.xlsx",sheet_name="FNMA LLPAs")
df_sec_col = df.iloc[:,1]
print (type(df_sec_col))
all_data_list = []
change_fico_data = False
main_col_head = ["FICO","Product Feature","LTV"]
col_fico_dict = {"fico_idx": [], "fico_cols": [[]]}
print("hello",df_sec_col)
counter = 0
version = 1
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

table_names = ["Fico15", "Cash Out Refinance", "Product Features"]
json_list = []
fico_dict ={
    "minFico" :0,
    "maxFico" :0
}
for outer_idx, fico_data in enumerate(all_data_list):
    json_out_dict = {
        "Product":'',
        "FICO": [],
        "LTV": []
    }
    for idx, data in enumerate(fico_data["fico_cols"]):
        if idx==0:
            json_out_dict["FICO"] = data[1:]
            for indx , info in enumerate(json_out_dict["FICO"]):
                # json_out_dict["FICO"][indx] = data_min+'-'+data_max
                print("here",json_out_dict["FICO"][indx])
                if 'Purchase (Fixed Rate)'  not in  data[1:]:
                    if '-' in info :
                        data_min, data_max = info.split('-')
                        fico_dict["minFico"] = int(data_min)
                        fico_dict["maxFico"] = int(data_max)
                        json_out_dict["FICO"][indx] = copy.deepcopy(fico_dict)
                    if '≥' in info :
                        data_max, data_min = info.split('≥')
                        data_max = str(sys.maxsize)
                        # json_out_dict["FICO"][indx] = data_min+'-'+data_max
                        # print("here",json_out_dict["FICO"][indx])
                        fico_dict["minFico"] = int(data_min)
                        fico_dict["maxFico"] = int(data_max)
                        json_out_dict["FICO"][indx] = copy.deepcopy(fico_dict)





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
            json_out_dict["Product"] = table_names[outer_idx]
            # json_list[0][table_names[outer_idx]] = copy.deepcopy(json_out_dict)
            json_list.append(copy.deepcopy(json_out_dict))
main_fico_list = []
fico_data_dict ={
    "Version":version,
    "product":"",
    "minFico":'',
    "maxFico":'',
    "llpas":[

        ]
}
product_feature_dict ={
"product":"",
    "productFeature":"",
    "llpas":[

        ]
}
llpa_dict = {
            "minLtv":23,
            "maxLtv":23,
            "llpaValue":23
        }
# for val in json_list :
#     for indx, ficoVal in enumerate(val['FICO']):
#         if val["Product"] !="Product Features" :
#             fico_data_dict["minFico"] = copy.deepcopy(ficoVal["minFico"])
#             fico_data_dict["maxFico"] = copy.deepcopy(ficoVal["maxFico"])
#             main_fico_list.append(copy.deepcopy(fico_data_dict))
#
#             for ltv_indx, ltvVal in enumerate(val["LTV"]):
#
#                 llpa_dict["minLtv"] = ltvVal["min"]
#                 llpa_dict["maxLtv"] = ltvVal["max"]
#                 # llpa_dict["llpaValue"] = ltvVal["values"][indx]
#                 # fico_data_dict["llpas"].append(copy.deepcopy(llpa_dict))
#                 fico_data_dict["llpas"].append(indx)
                #
                #
                # llpa_dict = {
                #     "minLtv": 23,
                #     "maxLtv": 23,
                #     "llpaValue": 23
                # }

            # for indx , info in enumerate(val["FICO"]):
            # json_out_dict["FICO"][indx] = data_min+'-'+data_max
            # print("here",json_out_dict["FICO"][indx])
            # if '-' in info :
            #     data_min, data_max = info.split('-')
            #     fico_dict_data["minFico"] = data_min
            #     fico_dict_data["maxFico"] = data_max
            #     json_out_dict["FICO"][indx] = copy.deepcopy(fico_dict)
            # if '≥' in info :
            #     data_max, data_min = info.split('≥')
            #     data_max = str(sys.maxsize)
            #     # json_out_dict["FICO"][indx] = data_min+'-'+data_max
            #     # print("here",json_out_dict["FICO"][indx])
            #     fico_dict["minFico"] = data_min
            #     fico_dict["maxFico"] = data_max
            #     json_out_dict["FICO"][indx] = copy.deepcopy(fico_dict)
productFeature = json_list.pop()
productFeature_list =[]
productFeature_list.append(productFeature)
is_product = False
for val in json_list:

    for index,stateVal in enumerate(val['FICO']):
        print(stateVal, index)
        for loanVal in val["LTV"]:
            print(loanVal)
            print(loanVal["values"][index])
            fico_data_dict["product"] = val["Product"]
            print("here what you looking for...!",stateVal)
            fico_data_dict["minFico"]= stateVal["minFico"]
            fico_data_dict["maxFico"]= stateVal["maxFico"]



            llpa_dict["minLtv"] = loanVal["min"]
            llpa_dict["maxLtv"] = loanVal["max"]
            llpa_dict["llpaValue"] = loanVal["values"][index]
            fico_data_dict["llpas"].append(llpa_dict)
            llpa_dict={

            }
        main_fico_list.append(copy.deepcopy(fico_data_dict))
        fico_data_dict={
            "Version": version,
            "product":"",
            "minFico":"",
            "maxFico":"",
            "llpas":[
            ]
}
print(productFeature_list)
for val in productFeature_list:

    for index, stateVal in enumerate(val['FICO']):
        print(stateVal, index)
        for loanVal in val["LTV"]:
            print(loanVal)
            print(loanVal["values"][index])
            product_feature_dict["product"] = val["Product"]
            print("here what you looking for...!", stateVal)
            product_feature_dict["productFeature"]=stateVal

            llpa_dict["minLtv"] = loanVal["min"]
            llpa_dict["maxLtv"] = loanVal["max"]
            llpa_dict["llpaValue"] = loanVal["values"][index]
            product_feature_dict["llpas"].append(llpa_dict)
            llpa_dict = {

            }
        main_fico_list.append(copy.deepcopy(product_feature_dict))
        product_feature_dict = {
            "Version":version,
            "product": "",
            "productFeature":"",
            "llpas": [
            ]
        }

print(fico_data_dict)
print(main_fico_list)
json_object = json.dumps(main_fico_list, indent=4)
print(json_object)


with open("json_dataFNMA_LLPAs.json","w") as j_file:
    print("Hello")
    j_file.write(json_object)
# print(json_list)
