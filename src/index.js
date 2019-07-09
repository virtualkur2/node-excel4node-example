const express = require('express');
const xl = require('excel4node');
const path = require('path');

let wb = new xl.Workbook({
  defaultFont: {
    size: 12,
    name: 'Calibri',
    color: '#FF0800'
  }
});

let ws = wb.addWorksheet('Hoja 1');

ws.cell(1,1).number(100);
ws.cell(1,2).number(200);
ws.cell(1,3).formula('A1 + B1');

const app = express();

app.use(express.json());
app.use(express.urlencoded({
  extended: true
}));

app.use((req, res, next) => {
  res.setHeader('Access-Control-Allow-Origin', 'localhost'); // <- Habilitado localhost para desarrollo
  res.setHeader('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept')
  next();
});

app.get('/file', (req, res) => {
  let xlFile = wb.write('MyFile.xlsx', res);
})

app.get('/', (req, res) => {
  let template = `
  <!DOCTYPE html>
  <html lang="en">
  <head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>Get File</title>
  </head>
  <body>
    <a href="http://localhost:3000/file">Descargar</a>
  </body>
  </html>
  `
  res.status(200).send(template);
});

app.listen(3000, () => {
  console.log('App started');
});
