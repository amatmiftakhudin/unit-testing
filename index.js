var Excel = require('exceljs');

var Client = require('node-rest-client').Client;

var client = new Client();

var workbook = new Excel.Workbook();



workbook.xlsx.readFile('api-list.xlsx')
	.then(function(){

		var apiListSheet = workbook.getWorksheet('Api List');
		var configLocal = workbook.getWorksheet('Local');

		//console.log(apiListSheet);

		var rowCount = apiListSheet.rowCount;
		for(i=1; i<rowCount; i++){
			let position = i+1;
			let path = apiListSheet.getCell('B' + position).value;
			let host = configLocal.getCell('B1').value.text;
			console.log(host);
			client.get(host + path, function(data, response){
				//console.log(data);
				apiListSheet.getCell('E' + position).value = data;
				console.log(apiListSheet.getCell('E' + position).value);
				workbook.xlsx.writeFile('api-list.xlsx');
			});
		}
		
		//console.log(rowCount);

		//console.log('File exist.');
	});

// client.get("https://klobid.free.beeceptor.com/rest/view/users", function(data, response){
// 	console.log(data);


// });