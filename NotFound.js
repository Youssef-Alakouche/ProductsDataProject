const ExcelJS = require("exceljs");

async function NotFoundedProductsFun(file) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(file);

  // Get the first worksheet
  const worksheet = workbook.getWorksheet(1);

  let NameRowIndex = 1; // default

  let DataModel = { Title: "" };
  let ExtractedData = [];

  let NotFounded = false;

  // Loop through rows and columns to extract data
  worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
    if (rowNumber == 1) {
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        // console.log(colNumber);
        if (cell.value.toLowerCase() == "title") {
          NameRowIndex = colNumber;
        }
      });
    } else {
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        // console.log(colNumber);
        if (colNumber == NameRowIndex) {
          DataModel.Title = cell.value;
          //   console.log(cell.value);
        }

        if (colNumber == 2) {
          //   NotFounded = true;
          //   console.log(cell.text == "");
          if (
            cell.value != null &&
            cell.value.toString().toLowerCase() == "not found"
          ) {
            NotFounded = true;
          }
        }
      });
      if (NotFounded) {
        ExtractedData.push({ ...DataModel });
        NotFounded = false;
      }
    }
    // console.log(`Row ${rowNumber}:`);
  });

  ExtractedData.reverse();
  // console.log(ExtractedData);

  return ExtractedData;
}
// let products = str.split("\n");

// console.log(products);

module.exports = { NotFoundedProductsFun };
