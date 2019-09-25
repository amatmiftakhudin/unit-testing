const dotenv = require('dotenv');
dotenv.config();

var Excel = require('exceljs');

var Client = require('node-rest-client').Client;

var client = new Client();

var workbook = new Excel.Workbook();

function compareObjects(object1, object2){
    var equal = true;
    for (i in object1)
        if (!object2.hasOwnProperty(i))
            equal = false;
    return equal;
}

const excelFile = process.env.EXCEL_FILE;

workbook.xlsx.readFile(excelFile)
	.then(function(){

		var apiListSheet = workbook.getWorksheet('Api List');
		var configLocal = workbook.getWorksheet('Local');

		//console.log(apiListSheet);

		var rowCount = apiListSheet.rowCount;
		for(i=1; i<rowCount; i++){
			let position = i+1;
			let path = apiListSheet.getCell('B' + position).value;
			let host = configLocal.getCell('B1').value.text;
			let method = apiListSheet.getCell('C' + position).value;
			let expectedResult = JSON.parse( apiListSheet.getCell('F' + position).value );
console.log(host + path)
			if(method == 'GET'){
				client.get(host + path, function(data, response){
					
					apiListSheet.getCell('E' + position).value = data;
					apiListSheet.getCell('G' + position).value = compareObjects(data, expectedResult) == true ? "PASS" : "FAILED";
					workbook.xlsx.writeFile(excelFile);
				});
			} else if (method == 'POST'){
				client.post(host + path, function(data, response){
					apiListSheet.getCell('E' + position).value = data;
					apiListSheet.getCell('G' + position).value = compareObjects(data, expectedResult) == true ? "PASS" : "FAILED";
					workbook.xlsx.writeFile(excelFile);
				});
			}

			
		}
		
		//console.log(rowCount);

		//console.log('File exist.');
	});

// client.get("https://klobid.free.beeceptor.com/rest/view/users", function(data, response){
// 	console.log(data);


// });