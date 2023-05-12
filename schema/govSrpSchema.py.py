gov_schema = {
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