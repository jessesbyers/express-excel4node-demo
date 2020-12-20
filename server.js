// gist from https://gist.github.com/natergj/62b9d2bfd3e2c6adf87a

var express = require('express');
var xl = require('excel4node');

var app = express();

app.get('/', function(req, res){
	res.end('Hello World');
});

app.get('/myreport', function(req, res) {
  makeReport(req, res);
});

var server = app.listen(3000, function () {
  console.log('Example app listening at http://%s:%s', '127.0.0.1', 3000);
});

function makeReport(req, res){
	var wb = new xl.Workbook();
	var ws = wb.addWorksheet('Sheet1');

    // ws.cell(1,1).String('String');
    // Create a reusable style
var style = wb.createStyle({
    font: {
      color: '#FF0800',
      size: 12,
    },
    numberFormat: '$#,##0.00; ($#,##0.00); -',
  });
   
  // Set value of cell A1 to 100 as a number type styled with paramaters of style
  ws.cell(1, 1)
    .number(100)
    .style(style);
   
  // Set value of cell B1 to 200 as a number type styled with paramaters of style
  ws.cell(1, 2)
    .number(200)
    .style(style);
   
  // Set value of cell C1 to a formula styled with paramaters of style
  ws.cell(1, 3)
    .formula('A1 + B1')
    .style(style);
   
  // Set value of cell A2 to 'string' styled with paramaters of style
  ws.cell(2, 1)
    .string('string')
    .style(style);
   
  // Set value of cell A3 to true as a boolean type styled with paramaters of style but with an adjustment to the font size.
  ws.cell(3, 1)
    .bool(true)
    .style(style)
    .style({font: {size: 14}});

	wb.write('MyWorkBook.xlsx', res);
}