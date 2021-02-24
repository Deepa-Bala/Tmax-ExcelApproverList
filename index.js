const express = require('express');
const fileUpload = require('express-fileupload');
const cors = require('cors');
const bodyParser = require('body-parser');
const morgan = require('morgan');
const _ = require('lodash');
const path = require('path');
var fs = require('fs');

const ExcelJS = require('exceljs');
const xlsxFile = require('read-excel-file');
const app = express();

const port = process.env.PORT || 3005;

app.get('/', function (req, res) {
    res.send('Hello World!');
  });
  
app.listen(port, () =>
    console.log(`App is listening on port ${port}.`)
);


app.post('/approver-identify', async (req, res) => {
    try {
        if (!req.files) {
            res.send({
                status: false,
                message: 'No file uploaded'
            });
        } else {
            let TemplateFile = req.files.TemplateFile;  
            await TemplateFile.mv('./uploads/' + TemplateFile.name);
            
                
            let filePath = './uploads/'+ TemplateFile.name;
            var workbook = new ExcelJS.Workbook();

            workbook.xlsx.readFile(filePath).then(function(){
            
                var worksheet = workbook.worksheets[0];  
                let approverlist =[
                    { "Level1" : worksheet.getRow(6).getCell(9).value} ,
                     {"Level2" : worksheet.getRow(7).getCell(9).value},
                     {"Level3" : worksheet.getRow(7).getCell(9).value},
                  ]            
                    console.log("Level1-"+worksheet.getRow(6).getCell(9).value);
                    console.log("Level2-"+worksheet.getRow(7).getCell(9).value);
                    console.log("Level3-"+worksheet.getRow(8).getCell(9).value);
                
                setTimeout(() => {
                        res.send(approverlist);            
                 }, 600);
                setTimeout(() => {      
                    fs.unlink(path.join(__dirname, "./uploads", path.normalize(`${TemplateFile.name}`)), ()=>{});
                    
                }, 1000);
            });
        }
        
    } catch (err) {
        res.status(500).send(err);
    }
});

app.post('/getapprover-identify', (req, res) => {
            let filePath = './uploads/Automated.xlsx';

            //let filePath = './uploads/'+ TemplateFile.name;
            var workbook = new ExcelJS.Workbook();

            workbook.xlsx.readFile(filePath).then(function(){
            
                var worksheet = workbook.worksheets[0];
                let approverlist =[
                   { "Level1" : worksheet.getRow(6).getCell(9).value} ,
                    {"Level2" : worksheet.getRow(7).getCell(9).value},
                    {"Level3" : worksheet.getRow(7).getCell(9).value},
                 ]      
                    console.log("Level1-"+worksheet.getRow(6).getCell(9).value);
                    console.log("Level2-"+worksheet.getRow(7).getCell(9).value);
                    console.log("Level3-"+worksheet.getRow(8).getCell(9).value);
                
                setTimeout(() => {
                        res.send(approverlist);            
                 }, 600);
                setTimeout(() => {      
                    //fs.unlink(path.join(__dirname, "./uploads", path.normalize(`${TemplateFile.name}`)), ()=>{});
                    
                }, 1000);
            });
});
