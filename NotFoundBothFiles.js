let { NotFoundedProductsFun } = require("./NotFound");
let ExcelJS = require("exceljs");

const CityChainFile = "./Search Products/citychainData.xlsx";
const thongsiaFile = "./Search Products/thongsiaData.xlsx";
const CommanNotFoundFile = "./Search Products/CommanNotFoundFile.xlsx";

getCommanNotFoundExcelFile();

async function getCommanNotFoundExcelFile() {
  let array1 = await NotFoundedProductsFun(thongsiaFile);
  let array2 = await NotFoundedProductsFun(CityChainFile);

  let commonNotFound = extractCommonObjects(array1, array2);

  const workbook = new ExcelJS.Workbook();
  const workSheet = workbook.addWorksheet("Scraped Data");

  workSheet.addRow(["Title"]);

  for (let product of commonNotFound) {
    let { Title } = product;
    workSheet.addRow([Title]);
  }

  await workbook.xlsx.writeFile(CommanNotFoundFile);
}

function extractCommonObjects(array1, array2) {
  const titleMap = new Map();

  // Build a map of titles from the first array
  for (const obj of array1) {
    titleMap.set(obj.Title, obj);
  }

  // Extract common objects from the second array
  const commonObjects = [];
  for (const obj of array2) {
    if (titleMap.has(obj.Title)) {
      commonObjects.push(titleMap.get(obj.Title));
      titleMap.delete(obj.Title); // Remove the entry to avoid duplication
    }
  }

  return commonObjects;
}
