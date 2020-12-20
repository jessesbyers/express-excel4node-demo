// gist from https://gist.github.com/natergj/62b9d2bfd3e2c6adf87a, but fixed multiple typos

var express = require('express');
var xl = require('excel4node');

var app = express();

app.get('/', function(req, res){
	res.end('Hello World');
});

app.get('/myreport', function(req, res) {
  makeReport(req, res);
});

var server = app.listen(8000, function () {
  console.log('Example app listening at http://%s:%s', '127.0.0.1', 8000);
});

function makeReport(req, res){
  // sample data
  const data = [{
    "id": 1,
    "first_name": "Jeanette",
    "last_name": "Penddreth",
    "email": "jpenddreth0@census.gov",
    "gender": "Female",
    "ip_address": "26.58.193.2"
  }, {
    "id": 2,
    "first_name": "Giavani",
    "last_name": "Frediani",
    "email": "gfrediani1@senate.gov",
    "gender": "Male",
    "ip_address": "229.179.4.212"
  }, {
    "id": 3,
    "first_name": "Noell",
    "last_name": "Bea",
    "email": "nbea2@imageshack.us",
    "gender": "Female",
    "ip_address": "180.66.162.255"
  }, {
    "id": 4,
    "first_name": "Willard",
    "last_name": "Valek",
    "email": "wvalek3@vk.com",
    "gender": "Male",
    "ip_address": "67.76.188.26"
  }]

  // set up workbook and worksheets
	var wb = new xl.Workbook();
  var ws1 = wb.addWorksheet('Sheet1');
  var ws2 = wb.addWorksheet('Sheet2');


    // Create reusable styles
  var headingStyle = wb.createStyle({
    font: {
      rowor: 'black',
      bold: true,
      size: 12,
    },
  });

  var bodyStyle = wb.createStyle({
    font: {
      rowor: '#FF0800',
      size: 12,
    },
  });
   
  // Set header rows; cell syntax is cell(row,column)
  // loop through keys and create a header for each column 
  // (i offset by 1 to match colums starting at 1)
  const headers = Object.keys(data[0])

  for (let i=1; i<headers.length+1; i++) {
    ws1.cell(1, i)
      .string(headers[i-1])
      .style(headingStyle);
  }



  //  Set values of each cell in each row by looping through data object. 
  // cell syntax is cell(row,column)
  for (let row=0; row<data.length; row++) {
    for (let col=0; col<headers.length; col++) {
      let header = headers[col]
      let value = data[row][header]
        // start at row #2 to be under headers; start at column #1
        ws1.cell(row+2, col+1)
          .string(value.toString())
          .style(bodyStyle)
    }
  }

  // set up worksheet 2

	wb.write('MyWorkBook.xlsx', res);
}