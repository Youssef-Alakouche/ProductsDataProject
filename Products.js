const ExcelJS = require("exceljs");

async function productsFun(file) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(file);

  // Get the first worksheet
  const worksheet = workbook.getWorksheet(1);

  let NameRowIndex = 1; // default

  let DataModel = { Title: "", IsScraped: false };
  let ExtractedData = [];

  //   this counter for extract unscrapted row
  let counter = 0;

  // Loop through rows and columns to extract data
  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    if (rowNumber == 1) {
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        // console.log(colNumber);
        if (cell.value != null && cell.value.toLowerCase() == "title") {
          NameRowIndex = colNumber;
        }
      });
    } else {
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        // console.log(colNumber);
        if (colNumber == NameRowIndex) {
          DataModel.Title = cell.value;
          // toString().replace(/\s*\(.*?\)\s*/g, "");
          //   console.log(cell.value);
        }

        if (colNumber > 1) {
          counter += 1;
          DataModel.IsScraped = true;
          //   console.log(cell.text == "");
        }
      });
      if (counter == 0) {
        DataModel.IsScraped = false;
      }

      ExtractedData.push({ ...DataModel });

      counter = 0;
    }
    // console.log(`Row ${rowNumber}:`);
  });

  ExtractedData.reverse();
  //   console.log(ExtractedData);

  return ExtractedData;
}
// let products = str.split("\n");

// console.log(products);

module.exports = { productsFun };
