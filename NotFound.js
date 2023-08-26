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

  // let counter = 0;

  // Loop through rows and columns to extract data
  worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
    if (rowNumber == 1) {
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        // console.log(colNumber);
        // console.log(colNumber + " " + cell.value);
        if (cell.value.toLowerCase() == "title") {
          NameRowIndex = colNumber;
        }
      });
    } else {
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        // console.log(colNumber);
        if (colNumber == NameRowIndex) {
          DataModel.Title = cell.value;
          //   console.log(cell.value);
        }

        // console.log(colNumber);

        if (colNumber == 2) {
          // NotFounded = false;

          // counter += 1;

          //   console.log(cell.text == "");
          if (
            cell.value != null &&
            cell.value.toString().toLowerCase() == "not found"
          ) {
            NotFounded = true;
          }
        }
      });
      // console.log("---------");
      if (NotFounded) {
        ExtractedData.push({ ...DataModel });
        // NotFounded = true;
      }
      NotFounded = true;
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
