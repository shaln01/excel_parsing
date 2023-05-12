from FNMA_LLPAs import fnma_llpa
from FHLMC_LLPAs import fhlmc_llpa
from ConvSRP_withEscrow import conv_srp_with_escrows
from ConvSRP_withoutEscrow import conv_srp_without_escrows
# from GovSRP_withEscrows import govSrp
# from GovSRP_LLPAs import  govSrpLlpa
# from PortfoliloExpress import portfolioExpress
# from PortfolioExpressLLPAs import portfolioExpressLlpa
# from PortfolioJumboLLPAs import  portfoliloJumboLlpa
# from PortfolioExpandedLLPAs import  portfolioExpandedLlpa
import  openpyxl
import pymongo
import time
import sys
from datetime import datetime

print(sys.argv[1])
version=(sys.argv[1])
version = int(version)
connection_string ="mongodb://localhost:27017"
sheet_path = "C:\\Users\\ShailendraSingh\\Desktop\\AprilSheet.xlsx"
wrkbk = openpyxl.load_workbook(sheet_path)
fnma_fhlmc_schema_validator ={
    "$jsonSchema": {
        # "_id" :"objectId",
                     "bsonType":'object',
                     "required":["version","isCashOutRefinance","llpas"],
        "properties": {
            "_id": {
                "bsonType":"objectId",
                "description": "must be objectid and is required"
            },
             "version": {
                 "bsonType":"int",
                   "description":"must be int and is required"
             },
            "provider": {
                "bsonType": "string",
                "description": "must be string and is required"
            },
            "isCashOutRefinance":{
                "bsonType":"bool",
                "description": "must be boolean and is required"

            },
            "product": {
                "bsonType": "string",
                "description": "must be string and is required"
            },

            "productFeature":{
            "enum": ["PURCHASE_FIXED","UNITS_2","UNITS_3_4" ,"INVESTMENT_PROPERTY" , "HOME_SECOND" ,"CONDO_ATTACHED_TERM_15_ABOVE" ,"ARM" ,"HIGHBAL_TERM_15_ABOVE","HIGHBAL_CASHOUT_YES" , "ARM_HIGHBAL" ,"TEMP_BUYDOWN" ],
            "description": "must be string and is required"
            },
            "llpas":{
                "bsonType": "array",
                # "minItems": 2,
                # "uniqueItems": True,
                "additionalProperties": False,
                 "items":{
                    "bsonType": "object",
                     # "required":["minLoan","maxLoan"],
                    "additionalProperties": True,
                    "description": "'items' must contain the stated fields.",
                    "properties": {
                        "minLtv": {
                          "bsonType": "number",
                          "description": "'minLoan' is required field of type number"
                                },
                        "maxLtv": {
                          "bsonType": "number",
                          "description": "'maxLoan' is a required field of type number"
                                },
                        "llpaValue": {
                            "bsonType": "number",
                            "description": "'srp' is a field of type number"
                        }
                    }
            }
            },
            "createdAt":{
                "bsonType": "date",
                "description": "must be a date and is required"
            },
            "updatedAt":{
                "bsonType": "date",
                "description": "must be a date and is required"
            }


        },

}
}
srp_validator = {
    "$jsonSchema": {
        # "_id" :"objectId",
                     "bsonType":'object',
                     "required":["version","withEscrow","product","state","term","productType","values"],
        "properties": {
            "_id": {
                "bsonType":"objectId",
                "description": "must be objectid and is required"
            },
             "version": {
                 "bsonType":"int",
                   "description":"must be int and is required"
             },
            "withEscrow":{
                "bsonType":"bool",
                "description": "must be boolean and is required"

            },
            "product": {
                "bsonType": "string",
                "description": "must be string and is required"
            },
            "state": {
                "bsonType": "string",
                "description": "must be string and is required"
            },
            "term": {
                "bsonType": "array",
                "description": "must be array and is required"
            },
            "productType":{
                "bsonType": "string",
                "description": "must be string and is required"
            },
            "values":{
                "bsonType": "array",
                "minItems": 2,
                "uniqueItems": True,
                "additionalProperties": False,
                 "items":{
                    "bsonType": "object",
                     "required":["minLoan","maxLoan"],
                    "additionalProperties": True,
                    "description": "'items' must contain the stated fields.",
                    "properties": {
                        "minLoan": {
                          "bsonType": "int",
                          "description": "'minLoan' is required field of type int"
                                },
                        "maxLoan": {
                          "bsonType": "number",
                          "description": "'maxLoan' is a required field of type number because the max number taken cannot be as int value"
                                },
                        "srp": {
                            "bsonType": "number",
                            "description": "'srp' is a field of type number"
                        }
                    }
            }
            },
            "createdAt":{
                "bsonType": "date",
                "description": "must be a date and is required"
            },
            "updatedAt":{
                "bsonType": "date",
                "description": "must be a date and is required"
            }
        },

}
}
conv_srp_validator ={
 "$jsonSchema": {
        # "_id" :"objectId",
                     "bsonType":'object',
                     "required":["version","withEscrow","product","state","term","productType","values"],
        "properties": {
            "_id": {
                "bsonType":"objectId",
                "description": "must be objectid and is required"
            },
             "version": {
                 "bsonType":"int",
                   "description":"must be int and is required"
             },
            "withEscrow":{
                "bsonType":"bool",
                "description": "must be boolean and is required"

            },
            "product": {
                "bsonType": "string",
                "description": "must be string and is required"
            },
            "state": {
                "bsonType": "string",
                "description": "must be string and is required"
            },
            "term": {
                "bsonType": "array",
                "description": "must be array and is required"
            },
            "productType":{
                "enum":["Fixed","ARM"],
                "description": "must be string and is required"
            },
            "values":{
                "bsonType": "array",
                "minItems": 2,
                "uniqueItems": True,
                "additionalProperties": False,
                 "items":{
                    "bsonType": "object",
                     "required":["minLoan","maxLoan"],
                    "additionalProperties": True,
                    "description": "'items' must contain the stated fields.",
                    "properties": {
                        "minLoan": {
                          "bsonType": "int",
                          "description": "'minLoan' is required field of type int"
                                },
                        "maxLoan": {
                          "bsonType": "number",
                          "description": "'maxLoan' is a required field of type number"
                                },
                        "srp": {
                            "bsonType": "number",
                            "description": "'srp' is a field of type number"
                        }
                    }
            }
            },
            "createdAt":{
                "bsonType": "date",
                "description": "must be a date and is required"
            },
            "updatedAt":{
                "bsonType": "date",
                "description": "must be a date and is required"
            }
        },

}
}
product_feature_validator ={
"$jsonSchema": {
    "bsonType": 'object',
    "required": ["version", "productFeature", "llpas"],
    "properties":{
        "version": {
            "bsonType": "int",
            "description": "must be int and is required"
        },
        "provider":{
            "bsonType": "string",
            "description": "must be string and is required"
        },
        "productFeature": {
            "enum":["PURCHASE_FIXED","UNITS_2","UNITS_3_4","INVESTMENT_PROPERTY","HOME_SECOND","CONDO_ATTACHED_TERM_15_ABOVE","ARM",
                    "HIGHBAL_TERM_15_ABOVE","HIGHBAL_CASHOUT_YES","ARM_HIGHBAL","TEMP_BUYDOWN"],
                "description": "must be string and is required"
        },
         "llpas":{
                "bsonType": "array",
                # "minItems": 2,
                # "uniqueItems": True,
                "additionalProperties": False,
                 "items":{
                    "bsonType": "object",
                     # "required":["minLoan","maxLoan"],
                    "additionalProperties": True,
                    "description": "'items' must contain the stated fields.",
                    "properties": {
                        "minLtv": {
                          "bsonType": "number",
                          "description": "'minLoan' is required field of type number"
                                },
                        "maxLtv": {
                          "bsonType": "number",
                          "description": "'maxLoan' is a required field of type number"
                                },
                        "llpaValue": {
                            "bsonType": "number",
                            "description": "'srp' is a field of type number"
                        }
                    }
            }
            },
            "createdAt":{
                "bsonType": "date",
                "description": "must be a date and is required"
            },
            "updatedAt":{
                "bsonType": "date",
                "description": "must be a date and is required"
            }
    }

}
}
gov_llpa_validator = {
    "$jsonSchema": {
        # "_id" :"objectId",
                     "bsonType":'object',
                     "required":["version","withEscrow","product","values"],
        "properties": {
            # "_id": {
            #     "bsonType":"objectId",
            #     "description": "must be objectid and is required"
            # },
             "version": {
                 "bsonType":"int",
                   "description":"must be int and is required"
             },
            "withEscrow":{
                "bsonType":"bool",
                "description": "must be boolean and is required"

            },
            "product": {
                "bsonType": "string",
                "description": "must be string and is required"
            },
            "state": {
                "bsonType": "string",
                "description": "must be string and is required"
            },
            "term": {
                "bsonType": "array",
                "description": "must be array and is required"
            },
            "productType":{
                "bsonType": "string",
                "description": "must be string and is required"
            },
            "values":{
                "bsonType": "array",
                "minItems": 2,
                # "uniqueItems": True,
                "additionalProperties": False,
                 "items":{
                    "bsonType": "object",
                     # "required":["minLoan","maxLoan"],
                    "additionalProperties": True,
                    "description": "'items' must contain the stated fields.",
                    "properties": {
                        "minLoan": {
                          "bsonType": "number",
                          "description": "'minLoan' is required field of type number"
                                },
                        "maxLoan": {
                          "bsonType": "number",
                          "description": "'maxLoan' is a required field of type number"
                                },
                        "minLtv": {
                          "bsonType": "number",
                          "description": "'minLoan' is field of type number"
                                },
                        "maxLtv": {
                          "bsonType": "number",
                          "description": "'maxLoan' is a  field of type number"
                                },
                        "minCltv": {
                          "bsonType": "number",
                          "description": "'minCltv' is a field of type number"
                                },
                        "maxCltv": {
                          "bsonType": "number",
                          "description": "'maxCltv' is a field of type number"
                                },

                        "productType":{
                            "bsonType":"string",
                            "description":"'productType' is a field type string"
                        },
                        "llpaValue": {
                            "bsonType": "number",
                            "description": "'srp' is a field of type number"
                        }
                    }
            }
            },
            "tier":{
                    "bsonType":"int",
                   "description":"'tier' is a field type int"
            },
            "stateGroup":{
                "bsonType": "array",
                "minItems": 1,
                "uniqueItems": True,
                "additionalProperties": False,
            },
            "createdAt":{
                "bsonType": "date",
                "description": "must be a date and is required"
            },
            "updatedAt":{
                "bsonType": "date",
                "description": "must be a date and is required"
            }


        },

}
}
if __name__ =="__main__":

    # fnma_llpa_data,fnma_llpa_product_feature_data = fnma_llpa(version,sheet_path)
    # fhlmc_llpa_data,fhlmc_llpa_product_feature_data = fhlmc_llpa(version,sheet_path)
    conv_srp_with_escrows_data = conv_srp_with_escrows(version,sheet_path)
    conv_srp_without_escrows_data = conv_srp_without_escrows(version,sheet_path)
    # gov_srp_data = govSrp(wrkbk,version)
    # gov_llpa_data = govSrpLlpa(wrkbk,version)
    # portfolio_express_data = portfolioExpress(wrkbk,version)
    # portfolio_express_llpa_data = portfolioExpressLlpa(wrkbk,version)
    # portfolio_jumbo_data = portfoliloJumboLlpa(wrkbk,version)
    # portfolio_expanded_data = portfolioExpandedLlpa(wrkbk,version)
    print("Welcome to pyMongo")
    client = pymongo.MongoClient(connection_string)
    # print(client)
    db = client["finance"]
    db_collections =db.list_collection_names()
    llpa_collection=''
    conv_srp_collection = ''
    product_feature_collection = ''
    # print(db.mySampleCollectionForShaln)
    # if db.mySampleCollectionForFHLMCLlpaValuesCollection is not None:
    #     db.mySampleCollectionForFHLMCLlpaValuesCollection.drop()
    # if db.mySampleCollectionForSrpValuesCollection is not None:
    #     print("Yes")
    #     db.mySampleCollectionForSrpValuesCollection.drop()
        # db.collection.remove('mySampleCollectionForShaln')
    # if db.mySampleCollectionForGovLlpaValuesCollection is not None:
    #     print("llpa Yes")
    #     db.mySampleCollectionForGovLlpaValuesCollection.drop()
    # if db.mySampleCollectionForFNMALlpaValuesCollection is not None:
    #     print("FNMA llpa Yes")
    #     db.mySampleCollectionForFNMALlpaValuesCollection.drop()
    # else:
    #     pass


    # srp_collection  = db.create_collection('mySampleCollectionForSrpValuesCollection',validator=srp_validator)
    # db.command("collection", "shaln", validator = srp_validator)
    # gov_llpa_collection  = db.create_collection('mySampleCollectionForGovLlpaValuesCollection',validator=gov_llpa_validator)
    # fhlmc_llpa_collection = db.create_collection('mySampleCollectionForFHLMCLlpaValuesCollection',validator =fnma_fhlmc_schema_validator)
    print("collection names")
    print(db.list_collection_names())
    # if 'llpaValues' not in db_collections:
    #     print("in none case")
    #     llpa_collection = db.create_collection('llpaValues', validator=fnma_fhlmc_schema_validator)
    # else:
    #     print("in not in none case")
    #     llpa_collection = db.llpaValues
    if "srpValues" not in db_collections:
        conv_srp_collection = db.create_collection('srpValues', validator=conv_srp_validator)

    else:
        conv_srp_collection = db.srpValues
    # if "productFeatureLlpas" not in db_collections:
    #     product_feature_collection = db.create_collection('productFeatureLlpas', validator=product_feature_validator)
    # else:
    #     product_feature_collection = db.productFeatureLlpas


    # for val in  gov_srp_data :
    #     val["createdAt"] = datetime.utcfromtimestamp(time.time())
    #     srp_collection.insert_one(val)
    with client.start_session() as session:
        with session.start_transaction():
            try:
                # for fnma_val in fnma_llpa_data:
                #     fnma_val["createdAt"] = datetime.utcfromtimestamp(time.time())
                #     fnma_val["updatedAt"] = datetime.utcfromtimestamp(time.time())
                #
                #     llpa_collection.insert_one(fnma_val, session=session)
                #
                #
                # for fnma_val in fnma_llpa_product_feature_data:
                #     fnma_val["createdAt"] = datetime.utcfromtimestamp(time.time())
                #     fnma_val["updatedAt"] = datetime.utcfromtimestamp(time.time())
                #
                #     product_feature_collection.insert_one(fnma_val,session=session)
                #
                #
                #
                # for fhlmc_val in fhlmc_llpa_data:
                #     fhlmc_val["createdAt"] = datetime.utcfromtimestamp(time.time())
                #     fhlmc_val["updatedAt"] = datetime.utcfromtimestamp(time.time())
                #
                #     llpa_collection.insert_one(fhlmc_val,session=session)
                #
                #
                #
                # for fhlmc_val in fhlmc_llpa_product_feature_data:
                #     fhlmc_val["createdAt"] = datetime.utcfromtimestamp(time.time())
                #     fhlmc_val["updatedAt"] = datetime.utcfromtimestamp(time.time())
                #
                #     product_feature_collection.insert_one(fhlmc_val,session=session)


                for srp_val in conv_srp_with_escrows_data:
                    srp_val["createdAt"] = datetime.utcfromtimestamp(time.time())
                    srp_val["updatedAt"] = datetime.utcfromtimestamp(time.time())
                    conv_srp_collection.insert_one(srp_val,session=session)


                # raise Exception("test exception")
                for srp_val in conv_srp_without_escrows_data:
                    srp_val["createdAt"] = datetime.utcfromtimestamp(time.time())
                    srp_val["updatedAt"] = datetime.utcfromtimestamp(time.time())
                    conv_srp_collection.insert_one(srp_val,session=session)


                session.commit_transaction()

            except Exception as e:
                print("Error- occured:",e)
                session.abort_transaction()
                # session.abort_transaction()





    #
    # for gov_llpa_val in gov_llpa_data:
    #     gov_llpa_val["createdAt"] = datetime.utcfromtimestamp(time.time())
    #
    #     gov_llpa_collection.insert_one(gov_llpa_val)

    # collection.insert_many(portfolio_express_data)
    # collection.insert_many(portfolio_express_llpa_data)
    # collection.insert_many(portfolio_jumbo_data)
    # collection.insert_many(portfolio_expanded_data)


# print(output_data)
