const { productsFun } = require("./Products");
const { NotFoundedProductsFun } = require("./NotFound");
const puppeteer = require("puppeteer");
const ExcelJS = require("exceljs");

const NOT_FOUND = "not found";
let ScrapedDataIndicator = false;

const SOURCE_FILE = "./Search Products/citychainData.xlsx";
const NOT_FOUND_FILE = "./Search Products/citychainNotFoundData.xlsx";

// let products = ["SRPE05K1", "SRPD85K1", "SRPD83K1", "SRPD81K1"];

let SearchedProductsArray = [];

function ResultObj(
  Title,
  Name = NOT_FOUND,
  Description = NOT_FOUND,
  Image = NOT_FOUND
) {
  this.Title = Title;
  this.Name = Name;
  this.Description = Description;
  this.Image = Image;
}

getExcelProduct();

async function getExcelProduct() {
  const browser = await puppeteer.launch({ headless: "new" });
  const page = await browser.newPage();

  console.log("start");
  // console.log(products);

  let products = await productsFun(SOURCE_FILE);
  // console.log(products);
  for (let { Title: productTitle, IsScraped } of products) {
    if (!IsScraped) {
      let PageUrl = getSearchingPage(productTitle);
      let result = await getExactProduct(page, PageUrl);

      if (result != NOT_FOUND) {
        let { Name, Description, image } = await ProductInfo(page, result);
        //   console.log(image);
        SearchedProductsArray.push(
          Object.assign(
            {},
            new ResultObj(productTitle, Name, Description, image)
          )
        );
      } else {
        SearchedProductsArray.push(
          Object.assign({}, new ResultObj(productTitle))
        );
      }
    } else {
      ScrapedDataIndicator = true;
    }
  }
  console.log("finish");

  SearchedProductsArray.reverse();
  await PopulatingExcelFile(SearchedProductsArray, ScrapedDataIndicator);

  await browser.close();
}

async function getExactProduct(page, PageUrl) {
  // const browser = await puppeteer.launch({ headless: "new" });
  // const page = await browser.newPage();
  // console.log("into getEx");
  await page.goto(PageUrl);

  //   i think you need extra check here,for the exact product key
  let [a] = await page.$x('//*[@id="content"]/div[2]/div/div/a[1]');

  let value;
  if (a != undefined) {
    const href = await a.getProperty("href");
    value = await href.jsonValue();
  } else {
    value = NOT_FOUND;
  }

  // await browser.close();
  return value;
}

async function ProductInfo(page, productPageUrl) {
  // const browser = await puppeteer.launch({ headless: "new" });
  // const page = await browser.newPage();
  // console.log("into prodinfo");
  await page.goto(productPageUrl);

  const [e1] = await page.$x(
    '//*[@id="content"]/div[2]/div/div[2]/div/div[2]/div[2]'
  );
  const Name = await page.evaluate((e) => e.textContent, e1);

  const [e2] = await page.$x(
    '//*[@id="content"]/div[2]/div/div[2]/div/div[2]/div[4]'
  );
  const Description = await page.evaluate((e) => e.textContent, e2);

  const [e3] = await page.$x(
    '//*[@id="content"]/div[2]/div/div[2]/div/div[1]/div[1]/img'
  );
  let src = await e3.getProperty("src");
  src = await src.jsonValue();

  return { Name, Description, image: src };

  // browser.close();
}

function getSearchingPage(ProductTitle) {
  let baseUrl = "https://www.thongsia.com.sg/en/seiko/search/?search=";

  return baseUrl + ProductTitle;
}

async function PopulatingExcelFile(SearchingProducts, indicator) {
  if (!indicator) {
    const workbook = new ExcelJS.Workbook();
    const workSheet = workbook.addWorksheet("Scraped Data");

    workSheet.addRow(["Title", "Name", "Description", "images"]);
    for (let product of SearchingProducts) {
      let { Title, Name, Description, Image } = product;
      console.log(Image);
      workSheet.addRow([Title, Name, Description, Image]);
    }

    await workbook.xlsx.writeFile(SOURCE_FILE);
  } else {
    console.log("lkjl");
    await updateExcelFile(SearchingProducts, SOURCE_FILE);
  }

  await NotFoundProductExcelFile(NOT_FOUND_FILE);
}

async function updateExcelFile(newProductsInfo, file) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(file);

  // Get the first worksheet
  const worksheet = workbook.getWorksheet(1);
  for (let product of newProductsInfo) {
    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      //
      if (rowNumber != 1) {
        row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
          if (colNumber == 1 && cell.value == product.Title) {
            console.log(product.Title);

            row.getCell(2).value = product.Name;
            row.getCell(3).value = product.Description;
            row.getCell(4).value = product.Image;

            row.commit();
          }
        });
      }
    });
  }

  await workbook.xlsx.writeFile(file);
}

async function NotFoundProductExcelFile(file, func = NotFoundedProductsFun) {
  const workbook = new ExcelJS.Workbook();
  const workSheet = workbook.addWorksheet("Scraped Data");

  workSheet.addRow(["Title"]);

  let notFoundProducts = await func(SOURCE_FILE);

  for (let product of notFoundProducts) {
    let { Title } = product;
    workSheet.addRow([Title]);
  }

  await workbook.xlsx.writeFile(file);
}
