var Excel = require('exceljs');
const assert = require("chai").assert;

const  workbook = new Excel.Workbook();
workbook.creator ="Naveen"; 
workbook.modified ="Kumar";

let data = [];

async function test(){
    await workbook.xlsx.readFile("Orders27112023.xlsx").then(function(){
        var workSheet =  workbook.getWorksheet("Sheet1"); 
    
        
    
        workSheet.eachRow({ includeEmpty: true }, function(row, rowNumber) {
            currRowArr = workSheet.getRow(rowNumber).values;
    
            currRow = workSheet.getRow(rowNumber); 
            // console.log("User Name :" + currRow.getCell(1).value +", Password :" +currRow.getCell(2).value);
            // console.log("User Name :" + row.values[1] +", Password :" +  row.values[2] ); 
            
    
            temp = currRowArr[2];
    
            data.push(currRowArr);
            console.log(currRowArr);
    
            //  assert.equal(currRow.getCell(2).type, Excel.ValueType.Number); 
           //  console.log("Row " + rowNumber + " = " + JSON.stringify(row.values));
        //    console.log(Object.values(currRowArr));
        //    console.log(currRowArr[2].hyperlink);
        //    console.log(currRowArr[20].hyperlink);
          });
    })
    
    getData();
}

async function getData(){

    await console.log(data);
}


test();