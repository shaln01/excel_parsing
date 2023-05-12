import pandas as pd
import numpy as np
import copy
import sys
import json
from enum import Enum
sheet_name = "FNMA LLPAs"
def fnma_llpa(version,sheet_path):
    df = pd.read_excel(sheet_path, sheet_name=sheet_name)
    df_sec_col = df.iloc[:, 1]
    all_data_list = []
    change_fico_data = False
    llpa_value_default = "null"
    col_fico_dict = {"fico_idx": [], "fico_cols": [[]]}
    counter = 0
    version = version
    max_fico_value =1000
    provider ='FHLMC'
    main_col_head = ["FICO", "Product Feature","LTV"]
    table_names = ["Purchase LLPAs (Terms > 15 years only)", "Purchase Loan Attribute LLPAs (all amortization terms)", "Rate/Term Refinance LLPAs (Terms > 15 years only)","Rate/Term Refinance Loan Attribute LLPAs (all amortization terms)","Cash Out Refinance LLPAs (all amortizatione terms)","Cash Out Loan Attribute LLPAs (all amortization terms)"]
    number_of_tables = len(table_names)
    print("running")
    class ProductEnum(Enum):
        PURCHASE_FIXED = ["Purchase (Fixed Rate)","Purchase (Fixed Rate)"]
        UNITS_2=["2 Units"]
        UNITS_3_4 =["3-4 Units"]
        INVESTMENT_PROPERTY = ["Investment Property"]
        HOME_SECOND = ["Second Homes"]
        CONDO_ATTACHED  = ["Attached  Condo","Condo  Attached","Attached Condo"]
        CONDO_ATTACHED_TERM_15_ABOVE = ["Attached  Condo & Term > 15yrs"]
        DTI_GREATER_THAN_40 = ["DTI > 40%"]
        LOANS_SECONDARY_FINANCING = ["Loans w/ Secondary Financing"]
        HIGHBAL_FRM = ["HighBal FRM"]
        HIGHBAL_ARM = ["HighBal ARM"]
        ARM = ["ARMs"]
        HIGHBAL_TERM_15_ABOVE = ["HighBal Purchase & No Cashout Refi (CLTV)"]
        HIGHBAL_CASHOUT_YES = ["HighBal Cashout Refi (CLTV)"]
        ARM_HIGHBAL = ["HighBal ARM (CLTV)"]
        TEMP_BUYDOWN = ["Loan with Temporary Buydown (Not subject to LLPA Caps)"]


    for row in df_sec_col:
        if row in main_col_head:
            if len(col_fico_dict["fico_idx"]) >= 1:
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
        if len(all_data_list) >= number_of_tables:
            break

    ltv_cols = df.iloc[:, 2:]

    for col in ltv_cols:
        for idx_range in all_data_list:
            col_data = list(df[col][idx_range["fico_idx"][0]:idx_range["fico_idx"][1]])
            idx_range["fico_cols"].append(col_data)

    # table_names = ["FICO  (Terms > 15 years only)", "Cash Out Refinance", "Product Features","Secondary Financing"]
    json_list = []
    fico_dict = {
        "minFico": 0,
        "maxFico": 0
    }
    for outer_idx, fico_data in enumerate(all_data_list):
        json_out_dict = {
            "Product": '',
            "FICO": [],
            "LTV": []
        }
        for idx, data in enumerate(fico_data["fico_cols"]):
            if idx == 0:
                json_out_dict["FICO"] = data[1:]
                for indx, info in enumerate(json_out_dict["FICO"]):

                    # print("here", json_out_dict["FICO"][indx])
                    if 'ARMs' not in data[1:]:
                        if '-' in info:
                            data_min, data_max = info.split('-')
                            fico_dict["minFico"] = float(data_min)
                            fico_dict["maxFico"] = float(data_max)
                            json_out_dict["FICO"][indx] = copy.deepcopy(fico_dict)
                        if '≥' in info:
                            data_max, data_min = info.split('≥')
                            data_max = max_fico_value
                            fico_dict["minFico"] = int(data_min)
                            fico_dict["maxFico"] = int(data_max)
                            json_out_dict["FICO"][indx] = copy.deepcopy(fico_dict)
                        if '<' in info:
                            data_min, data_max = info.split('<')

                            fico_dict["minFico"] = 0
                            fico_dict["maxFico"] = int(data_max)-1
                            json_out_dict["FICO"][indx] = copy.deepcopy(fico_dict)
                        if '<' in info:
                            data_min, data_max = info.split('<')

                            fico_dict["minFico"] = 0
                            fico_dict["maxFico"] = int(data_max)-1
                            json_out_dict["FICO"][indx] = copy.deepcopy(fico_dict)
                        if '≤' in info:
                            data_min, data_max = info.split('≤')

                            fico_dict["minFico"] = 0
                            fico_dict["maxFico"] = int(data_max)
                            json_out_dict["FICO"][indx] = copy.deepcopy(fico_dict)
                continue
            ltv_values = {
                "min": 0,
                "max": 0,
                "values": []
            }
            ltv_head = data[0]
            if "-" in str(ltv_head):
                # print("greater")

                ltv_min, ltv_max = str(ltv_head).split("-")

            elif "≥" in str(ltv_head):
                # print("greater",ltv_head[1:])
                ltv_min = ltv_head[1:]
                ltv_max = 1000
            elif "<" in str(ltv_head):
                ltv_max = int(ltv_head[1:])-1
                ltv_min = 0
            elif "CLTV" in str(ltv_head):
                ltv_max = 0
                ltv_min = 0
            else:
                ltv_min = 0
                ltv_max = ""
                for letter in str(ltv_head):
                    if letter.isnumeric():
                        ltv_max += letter

            ltv_values["min"] = ltv_min
            ltv_values["max"] = ltv_max
            ltv_val_df = pd.DataFrame(data[1:])
            ltv_val_df.fillna(llpa_value_default, inplace=True)
            # print(ltv_val_df[0].to_list())
            ltv_values["values"] = ltv_val_df[0].to_list()
            json_out_dict["LTV"].append(copy.deepcopy(ltv_values))
            if idx == len(fico_data["fico_cols"]) - 1:
                print(outer_idx)
                json_out_dict["Product"] = table_names[outer_idx]

                json_list.append(copy.deepcopy(json_out_dict))
    main_fico_list = []
    product_feature_list= []
    fico_data_dict = {
        "version": version,
        "provider":provider,
        "isCashOutRefinance":False,
        "product": "",
        "minFico": '',
        "maxFico": '',
        "llpas": [

        ]
    }


    product_feature_dict = {
        "version": version,
        "provider": provider,
        "productFeature": "",
        "llpas": [

        ]
    }
    llpa_dict = {
        "minLtv": 23,
        "maxLtv": 23,
        "llpaValue": 2
    }
    print("JSON LIST",json_list)
    print("JSON LIST LENGTH",len(json_list))

    # secondaryFinance =json_list.pop()
    secondaryFinance = []
    productFeature1 = json_list.pop(1)
    productFeature2=json_list.pop(2)
    productFeature3=json_list.pop(3)
    secondaryFinance_list =[]
    productFeature_list = []
    secondaryFinance_list.append(secondaryFinance)
    productFeature_list.append(productFeature1)
    # productFeature_list.append(productFeature2)
    # productFeature_list.append(productFeature3)

    is_product = False
    for val in json_list:

        for index, stateVal in enumerate(val['FICO']):
            # print(stateVal, index,"here is index")
            for loanVal in val["LTV"]:
                # print(loanVal)
                # print(loanVal["values"][index])
                if val["Product"]== "Purchase LLPAs (Terms > 15 years only)":
                    fico_data_dict["product"] = "FICO  (Terms > 15 years only)"
                else:
                    fico_data_dict["product"] = val["Product"]
                # print("here what you looking for...!", stateVal)
                fico_data_dict["minFico"] = stateVal["minFico"]
                fico_data_dict["maxFico"] = stateVal["maxFico"]
                if fico_data_dict["product"] == "Cash Out Refinance" or fico_data_dict["product"] == "Cash Out Refinance LLPAs (all amortizatione terms)":
                    fico_data_dict["isCashOutRefinance"] = True
                llpa_dict["minLtv"] = float(loanVal["min"])
                llpa_dict["maxLtv"] = float(loanVal["max"])
                if loanVal["values"][index] !=llpa_value_default:
                    llpa_dict["llpaValue"] = loanVal["values"][index]

                else:
                    pass
                fico_data_dict["llpas"].append(llpa_dict)

                llpa_dict = {

                }
            main_fico_list.append(copy.deepcopy(fico_data_dict))
            fico_data_dict = {
                "version": version,
                "provider": provider,
                "isCashOutRefinance": False,
                "product": "",
                "minFico": "",
                "maxFico": "",
                "llpas": [
                ]
            }
    llpa_dict = {
        "minLtv": 23,
        "maxLtv": 23,
        "llpaValue": 2
    }
    # print(productFeature_list)
    for val in productFeature_list:

        for index, stateVal in enumerate(val['FICO']):
            for loanVal in val["LTV"]:
                for prodVal in ProductEnum:
                    if stateVal in prodVal.value:
                        product_feature_dict["productFeature"] = prodVal.name
                llpa_dict["minLtv"] = float(loanVal["min"])
                llpa_dict["maxLtv"] = float(loanVal["max"])
                if loanVal["values"][index] != llpa_value_default:
                    llpa_dict["llpaValue"] = loanVal["values"][index]

                else:
                    pass
                product_feature_dict["llpas"].append(llpa_dict)


                llpa_dict = {

                }
            product_feature_list.append(copy.deepcopy(product_feature_dict))
            product_feature_dict = {
                "version": version,
                "provider": provider,
                "productFeature": "",
                "llpas": [
                ]
            }
    secondary_data_dict = {
        "version": version,
        "provider": provider,
        "isCashOutRefinance": False,
        "product": "",
        "minFico": '',
        "maxFico": '',
        "llpas": [

        ]
    }
    llpa_dict = {
        "minLtv": 23,
        "maxLtv": 23,
        "llpaValue": 23
    }

    val=secondaryFinance
    print("BEKAAAR HAI BHAIYA")
    json_object = json.dumps(main_fico_list, indent=4)

    json_product_object = json.dumps(product_feature_list, indent=4)

    with open("json_dataFNMA_LLPAsNewApril.json", "w") as j_file:
        # print("Hello")
        j_file.write(json_object)
    with open ("FNMA_productFeatureApril.json","w") as jp_file:

        jp_file.write(json_product_object)

    return (main_fico_list,product_feature_list)



