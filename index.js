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

main();

async function main() {

    // Open browser and website
    let driver = await new Builder().forBrowser("chrome").build();
    await driver.get("https://online.immi.gov.au/ola/app");


    await checkStatus();

    async function checkStatus() {
        for(let i = 2; i < names.length; i++) {
            let statusCell = 'G' + (i+1);
            let lastUpdateDateCell = 'H' + (i+1);

            
            if (wb.worksheets[0].getCell(statusCell).value == "Finalised" ) {

           } else {

            // Login
            await driver.findElement(By.name("username")).sendKeys(names[i]);
            await driver.findElement(By.name("password")).sendKeys(pass[i], Key.RETURN);
            await driver.findElement(By.name("continue")).sendKeys(Key.RETURN);

            // Copy Application Status from Website 
            let status = await driver.findElement(By.xpath("/html/body/form/section/div/div/div[3]/div/div[2]/div/div/div[2]/div/div[1]/div/div[2]/div/p")).getAttribute("innerText");
            let lastUpdateDate = await driver.findElement(By.xpath("/html/body/form/section/div/div/div[3]/div/div[2]/div/div/div[2]/div/div[2]/div/div/div[1]/div/div/div/div/div[2]/div/div/div[1]/div/div/div/div/time")).getAttribute("innerText");
            if (status == "Finalised") {
                console.log(names[i] + " is granted!")
            }
            
            // Paste Application Status from Website to Status Column in Excel file
            wb.worksheets[0].getCell(statusCell).value = status;
            wb.worksheets[0].getCell(lastUpdateDateCell).value = lastUpdateDate;
            console.log(names[i], lastUpdateDate);
            await wb.xlsx.writeFile('Immi Credentials.xlsx');
            
            
            // Logout
            await driver.findElement(By.xpath("/html/body/form/header/div/div/ol/li[3]/button")).click();
            await driver.findElement(By.xpath("/html/body/header/div/ul/li/div/a")).click();

            // Quit driver if no more accounts left to check
            if (i == names.length - 1) {
                await driver.quit();
            }
            }

            }

        }
        
}




