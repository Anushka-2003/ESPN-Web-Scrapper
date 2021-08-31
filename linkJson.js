const fs = require("fs");

const xlsx = require("xlsx");
// let buffer = fs.readFileSync("./example.json");
// // console.log(buffer);
// console.log("``````````````````");
// let data = JSON.parse(buffer);
// console.log(data);

let data = require("./example.json");
// console.log(data);

// data.push({
//     "name": "New",
//     "last Name": "Rogers",
//     "isAvenger": true,
//     "friends": [
//         "Bruce",
//         "Neter",
//         "Natasha"
//     ],
//     "age": 45,
//     "address": {
//         "city": "New York",
//         "state": "Manhatten"
//     }
// });

// let stringData = JSON.stringify(data);
// fs.writeFileSync("example.json", stringData);

// new worksheet

function excelWriter(filePath, json, sheetName){
    let newWB = xlsx.utils.book_new();

    // json data -> excel format convert
    let newWS = xlsx.utils.json_to_sheet(json);

    // add ws to wb
    xlsx.utils.book_append_sheet(newWB, newWS, sheetName);
    
    // to write in file system
    xlsx.writeFile(newWB, filePath);
}



function excelReader(filePath, sheetName){
    if(fs.existsSync(filePath) == false){
        return [];
    }

    // workbook get
    let wb = xlsx.readFile(filePath);
    
    // sheet 
    let excelData = wb.Sheets[sheetName];
    
    // sheet data get
    let ans = xlsx.utils.sheet_to_json(excelData);
    return;
    // console.log(ans);

}

