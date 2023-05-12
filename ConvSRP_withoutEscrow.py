import pandas as  pd
import numpy as np
import copy
import sys
import json
sheet_name ="Conv SRP - Without Escrows"
def conv_srp_without_escrows(version,sheet_path):
    df = pd.read_excel(sheet_path,sheet_name=sheet_name)

    # df = pd.read_excel("C:\\Users\\mahes\\OneDrive\\Desktop\\ameri.xlsx",sheet_name="Conv SRP - With Escrows")
    df_sec_col = df.iloc[:,1]
    # print (type(df_sec_col))
    all_data_list = []
    change_state_data = False
    main_col_head = ["State"]
    col_state_dict = {"state_idx": [], "state_cols": [[]]}
    # print("hello",df_sec_col)
    counter = 0
    counter = 0
    version = version
    provider = "Conv SRP"
    withEscrow =False
    table_names = ["Conventional 20/25/30 Year Fixed", "Conventional 10/15 Year Fixed", "Conventional 10/6 ARM","Conventional 7/6 ARM","Conventional 5/6 ARM"]
    number_of_tables = len(table_names)

    for row in df_sec_col:
        if row in main_col_head:
            if len(col_state_dict["state_idx"])>=1:
                col_state_dict["state_idx"].append(counter)
                all_data_list.append(copy.deepcopy(col_state_dict))
                col_state_dict = {"state_idx": [], "state_cols": [[]]}
            change_state_data = True
            col_state_dict["state_idx"].append(counter)
            col_state_dict["state_cols"][0].append(row)
        if row not in main_col_head and row is not np.nan and change_state_data:
            col_state_dict["state_cols"][0].append(row)
        if change_state_data and row is np.nan:
            col_state_dict["state_idx"].append(counter)
            all_data_list.append(copy.deepcopy(col_state_dict))
            col_state_dict = {"state_idx": [], "state_cols": [[]]}
            change_state_data = False
        counter += 1
        if len(all_data_list)>=number_of_tables:
            break

    ltv_cols = df.iloc[:,2:]

    for col in ltv_cols:
        for idx_range in all_data_list:
            col_data = list(df[col][idx_range["state_idx"][0]:idx_range["state_idx"][1]])
            idx_range["state_cols"].append(col_data)
    json_list = []
    main_json_list = []
    for outer_idx, state_data in enumerate(all_data_list):
        json_out_dict = {
            "Product":"",
            "State": [],
            "Loan Amount": []
        }
        # json_out_dict["Product"] = table_names[outer_idx]
        for idx, data in enumerate(state_data["state_cols"]):
            if idx==0:
                json_out_dict["State"] = data[1:]
                continue
            ltv_values = {
                "min": "",
                "max": "",
                "values": []
            }
            ltv_head = data[0]

            if "-" in ltv_head:
                ltv_min,ltv_max = ltv_head.split("-")
            elif ">" in ltv_head:
                ltv_max=str(sys.maxsize)
                ltv_min=ltv_head[1:]
            else:
                ltv_min = "0"
                ltv_max = ""
                for letter in ltv_head:
                    if letter.isnumeric():
                        ltv_max += letter
                        print(type(ltv_max))
            ltv_min = ltv_min.replace(',', '')
            ltv_max = ltv_max.replace(',', '')
            # print(type (ltv_max))

            ltv_values["min"] = int(ltv_min)
            ltv_values["max"] = int(ltv_max)
            ltv_val_df = pd.DataFrame(data[1:]) # written this just to get the series
            ltv_val_df.fillna("0",inplace=True) # trying to replace none
            # print(ltv_val_df[0].to_list())
            ltv_values["values"] = ltv_val_df[0].to_list()
            json_out_dict["Loan Amount"].append(copy.deepcopy(ltv_values))
            if idx==len(state_data["state_cols"])-1:
                # print(outer_idx)
                json_out_dict["Product"] = table_names[outer_idx]
                # print(outer_idx,json_out_dict)
                # json_list[0][table_names[outer_idx]] = copy.deepcopy(json_out_dict)
                json_list.append(copy.deepcopy(json_out_dict))

    json_list_obj ={
        "version": version,
        "withEscrow": withEscrow,
        "product":"",
        "state":"",
        "term":[],
        "productType":"",
        "values":[
        ]
    }
    data_value={
                "minLoan":122,
                "maxLoan":32,
                "srp": 12
            }
    for val in json_list:
        for index,stateVal in enumerate(val['State']):
            # print(stateVal, index)
            for loanVal in val["Loan Amount"]:
                # print(loanVal)
                # print(loanVal["values"][index])
                json_list_obj["product"] = val["Product"]
                json_list_obj["state"]= stateVal
                prod_packing=val["Product"].split(" ")
                terms=prod_packing[1]
                term_val = terms.split("/")
                # print('here is what you looking for',term_val)
                json_list_obj["term"]= term_val
                json_list_obj["term"]  = list(map(int, json_list_obj["term"]))
                json_list_obj["productType"] = prod_packing[-1]
                data_value["minLoan"] = loanVal["min"]
                data_value["maxLoan"] = loanVal["max"]
                data_value["srp"] = loanVal["values"][index]
                json_list_obj["values"].append(data_value)
                data_value={
                }
            main_json_list.append(copy.deepcopy(json_list_obj))
            json_list_obj={
        "version": version,
        "withEscrow": withEscrow,
        "product":"",
        "term": [],
        "productType": "",
        "state":"",
        "values":[
        ]
    }
    print('here is what u looking for the length of the array ',len(main_json_list))

    json_object = json.dumps(main_json_list, indent=4)
    with open("convSRP_withoutEscrows.json","w") as j_file:
        print("Hello")
        j_file.write(json_object)
    return (main_json_list)
