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

function jsonToQueryString(json) {
    return '?' + 
        Object.keys(json).map(function(key) {
            return encodeURIComponent(key) + '=' +
                encodeURIComponent(json[key]);
        }).join('&');
}

const excelFile = process.env.EXCEL_FILE;
var rootPathApi = process.env.ROOT_PATH;
var authToken = process.env.AUTH_TOKEN;
var showConsole = false;
if(process.env.ENV != null || process.env.ENV != undefined){
	rootPathApi = process.env["ROOT_PATH_" + process.env.ENV];
}

if(process.env.SHOW_CONSOLE != null || process.env.SHOW_CONSOLE != undefined){
	showConsole = ( process.env.SHOW_CONSOLE == 'true' || process.env.SHOW_CONSOLE == 'TRUE');
}
workbook.xlsx.readFile(excelFile)
	.then(function(){

		var apiListSheet = workbook.getWorksheet('Api List');
		var configLocal = workbook.getWorksheet('Local');

		var rowCount = apiListSheet.rowCount;
		for(i=1; i<rowCount; i++){
			let position = i+1;
			let path = apiListSheet.getCell('B' + position).value;
			let host = configLocal.getCell('B1').value.text;
			let method = apiListSheet.getCell('C' + position).value;
			let expectedResult = JSON.parse( apiListSheet.getCell('F' + position).value );
			let parameters = JSON.parse( apiListSheet.getCell('D' + position).value );
			var args = {
				data: parameters,
				headers: {
					"Content-Type": "application/json",
					"Authorization": authToken
				}
			};
			
			if(method == 'GET'){
				let qstr = jsonToQueryString(args.data);
				client.get(rootPathApi + path + qstr, args, function(data, response){
					if(showConsole == true){
						logData = JSON.stringify({url: rootPathApi + path, parameters: parameters, result: data}, null, "    ");
						console.log(logData);
					}

					apiListSheet.getCell('E' + position).value = data;
					apiListSheet.getCell('G' + position).value = compareObjects(data, expectedResult) == true ? "PASS" : "FAILED";
					workbook.xlsx.writeFile(excelFile);
				});
			} else if (method == 'POST'){
				client.post(rootPathApi + path, args, function(data, response){
					if(showConsole == true){
						logData = JSON.stringify({url: rootPathApi + path, parameters: parameters, result: data}, null, "    ");
						console.log(logData);
					}
					
					apiListSheet.getCell('E' + position).value = data;
					apiListSheet.getCell('G' + position).value = compareObjects(data, expectedResult) == true ? "PASS" : "FAILED";
					workbook.xlsx.writeFile(excelFile);
				});
			}

			
		}
	});