var MongoClient = require('mongodb').MongoClient;
var assert = require('assert');
var ObjectID = require('mongodb').ObjectID;
var utilityFunctions=require('./utilityFunctions');

//connecting to excel sheet to retrieve data
XLSX = require('xlsx');
var excel_file = process.argv[2];
var workbook = XLSX.readFile(excel_file);
var url = "mongodb://127.0.0.1:27017/eposroDB";

MongoClient.connect(url, function (err, db) {
    if (err) {
        console.log(err);
        return db.close();
    }
    //else part
    console.log("Connected Correctly to " + url);
    var IDs = db.collection("IDs");
    var Category = db.collection('Category');

    IDs.findOne(function (err, docs) {
        if (err) {
            console.log(err);
            return db.close();
        }
        var count = docs.last_entry_id;

        for (i = 1; i < workbook.SheetNames.length; i++) {
            current_sheet = workbook.SheetNames[i];
            worksheet = workbook.Sheets[current_sheet];
            data = XLSX.utils.sheet_to_json(worksheet);

            if (JSON.stringify(data) === '[]') {
                break;
            } else {
                var tcount = 1;
                for (var j = 0; j < data.length; j++) {
                    if (data[j]._id === '-') {
                        data[j]._id = ++count;
                        worksheet['A' + (tcount + 1)].v = count;
                    }
                    tcount += 1;
                }
                var product=utilityFunctions.get_to_product_format(data[i]);
                return db.close();
                //insert_to_database(product);
            }

        }
        XLSX.writeFile(workbook, "out.xlsx");
        IDs.update({
                _id: new ObjectID(docs._id)
            }, {
                "$set": {
                    last_entry_id: count
                }
            },
            function (err, nupdate) {
                if (err) {
                    console.error(err);
                    return db.close();
                } 
                return db.close();
            }
        );

    });

});