const express = require('express');
const ExcelJS = require('exceljs');

const app = express();
const port = 3001;

app.get('/read-excel', (req, res) => {
  const filePath = 'Chartdata.xlsx';

  const workbook = new ExcelJS.Workbook();

  workbook.xlsx
    .readFile(filePath)
    .then(() => {
      const worksheet = workbook.getWorksheet(1); // Get the first worksheet
      const data = [];

      worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
        // Access each row and its cell values
        data.push({ rowNumber, values: row.values });
      });

      res.json({ excelData: data });
    })
    .catch(error => {
      console.error('Error reading the file:', error);
      res.status(500).json({ error: 'Error reading the Excel file' });
    });
});

app.listen(port, () => {
  console.log(`Server running on http://localhost:${port}`);
});
