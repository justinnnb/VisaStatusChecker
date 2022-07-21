const {Builder, By, Key, until} = require('selenium-webdriver');

const ExcelJS = require('exceljs');
const { stat } = require('selenium-webdriver/io');

const sheet = './Immi Credentials/.xlsx';
const wb = new ExcelJS.Workbook();

let row;
let names = []
let pass = []
let statusColumn;

wb.xlsx.readFile('Immi Credentials.xlsx').then(() => { 
    row = wb.worksheets[0].getRow(4);
    usernameColumn = wb.worksheets[0].getColumn(4);
    passwordColumn = wb.worksheets[0].getColumn(5);
    statusColumn = wb.worksheets[0].getColumn(7);

    usernameColumn.eachCell( function(cell, rowNumber) {
        names.push(cell.value);
    })

    passwordColumn.eachCell( function(cell, rowNumber) {
        pass.push(cell.value);
    }

    
  )})
  console.log(names)
  console.log(pass)
