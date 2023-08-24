const { productsFun } = require("./Products");
const { NotFoundedProductsFun } = require("./NotFound");
// const { urls, ProductInfo} = require("./scraping");
const puppeteer = require("puppeteer");
const ExcelJS = require("exceljs");

const NOT_FOUND = "not found";
// indicat that there is Scraped data on the file
let ScrapedDataIndicator = false;

// let array = [];

const SOURCE_FILE = "./Search Products/thongsiaData.xlsx";
const NOT_FOUND_FILE = "./Search Products/thongsiaNotFoundData.xlsx";

// let products = [
//   "SWR078P1",
//   "SWR073P1",
//   "SWR035P1",
//   "SWR033P1",
//   "SUT405P1",
//   "SUT403P1",
//   "SUR634P1",
//   "SUR633P1",
//   "SUR632P1",
//   "SUR558P1",
// ];

let SearchedProductsArray = [];

function ResultObj(
  Title,
  Name = NOT_FOUND,
  Description = NOT_FOUND,
  Images = [NOT_FOUND]
) {
  this.Title = Title;
  this.Name = Name;
  this.Description = Description;
  this.Images = Images;
}

getExcelFile();

async function getExcelFile() {
  const browser = await puppeteer.launch({ headless: "new" });
  const page = await browser.newPage();

  console.log("started");
  let counter = 0;

  let products = await productsFun(SOURCE_FILE);
  products.reverse();

  for (let { Title: productTitle, IsScraped } of products) {
    // console.log(productName);

    if (!IsScraped) {
      let PageUrl = urls(undefined, productTitle)[0];
      let result = await resultNbr(page, PageUrl);

      // console.log(result);

      console.log(++counter);
      console.log(productTitle);

      if (result == "not Founded") {
        // const { Title, Description, Images } = new ResultObj(productName);
        SearchedProductsArray.push(
          Object.assign({}, new ResultObj(productTitle))
        );
        // workSheet.addRow([productName, Title, Description, ...Images]);
        // array.push(new ResultObj(productName));
      } else {
        const productPageUrl = await getExactProduct(page, PageUrl, result);

        if (productPageUrl == NOT_FOUND) {
          SearchedProductsArray.push(
            Object.assign({}, new ResultObj(productTitle))
          );
        } else {
          const productInfo = await ProductInfo(page, productPageUrl);
          // console.log(productInfo);

          let { Name, Specs: Description, images } = productInfo;

          // workSheet.addRow([productName, Title, Description, ...images]);

          SearchedProductsArray.push(
            Object.assign(
              {},
              new ResultObj(productTitle, Name, Description, images)
            )
          );
        }
      }
    } else {
      ScrapedDataIndicator = true;
    }
  }

  console.log("finished");

  // console.log(SearchedProductsArray);
  // await workbook.xlsx.writeFile("FirstWebsiteData.xlsx");
  // console.log(array);

  await PopulateExcelFile(
    SearchedProductsArray.reverse(),
    ScrapedDataIndicator
  );

  await browser.close();
}

async function getExactProduct(page, PageUrl, productsNbr) {
  // const browser = await puppeteer.launch({ headless: "new" });
  // const page = await browser.newPage();
  await page.goto(PageUrl);

  // let [e] = await page.$x(
  //   '//*[@id="CollectionAjaxContent"]/div[2]/div[1]/div[2]/div[1]'
  // );
  let [e] = await page.$$(
    "#CollectionAjaxContent > div.grid__item.medium-up--four-fifths.grid__item--content > div > div.grid.grid--uniform > div.grid-product__has-quick-shop"
  );
  // '//*[@id="CollectionAjaxContent"]/div[2]/div/div[2]/div[1]'
  // '#CollectionAjaxContent > div.grid__item.medium-up--four-fifths.grid__item--content > div > div.grid.grid--uniform > div:nth-child(1)'
  // if (productsNbr == 1) {
  //   [e] = await page.$$(
  //     "#CollectionAjaxContent > div.grid__item.medium-up--four-fifths.grid__item--content > div > div.grid.grid--uniform > div.grid-product__has-quick-shop"
  //   );
  //   // if (e == undefined) {
  //   //   [e] = await page.$x(
  //   //     '//*[@id="CollectionAjaxContent"]/div[2]/div/div[2]/div'
  //   //   );
  //   // }
  // }

  if (e == undefined) {
    return NOT_FOUND;
  }

  ("#CollectionAjaxContent > div.grid__item.medium-up--four-fifths.grid__item--content > div > div.grid.grid--uniform > div > div.grid-product__content > a");

  // '//*[@id="CollectionAjaxContent"]/div[2]/div/div[2]/div/div[1]/a'
  // '//*[@id="CollectionAjaxContent"]/div[2]/div/div[2]/div/div[1]/a'
  //   "#CollectionAjaxContent > div.grid__item.medium-up--four-fifths.grid__item--content > div > div.grid.grid--uniform > div"
  // "#CollectionAjaxContent > div.grid__item.medium-up--four-fifths.grid__item--content > div > div.grid.grid--uniform > div"
  // console.log(e);
  const anchors = await e.$$(
    "div.grid-product__content > a.grid-product__link"
  );
  console.log(anchors);
  const href = await anchors[0].getProperty("href");
  const value = await href.jsonValue();

  // console.log(value);

  // await browser.close();
  return value;
}

async function resultNbr(page, PageUrl) {
  // const browser = await puppeteer.launch({ headless: "new" });
  // const page = await browser.newPage();
  await page.goto(PageUrl);

  // const [element] = await page.$$(
  //   "div.collection-filter__item.collection-filter__item--count"
  // );

  const [element] = await page.$x(
    '//*[@id="CollectionAjaxContent"]/div[2]/div[1]/div[1]/div[2]'
  );
  let result;
  if (element != undefined) {
    result = await page.evaluate((e) => e.textContent, element);
    result = parseInt(result.split(" ")[0]);
  } else {
    result = "not Founded";
  }

  // console.log(result);

  // await browser.close();
  return result;
}

function urls(maxPages = 20, term = "sku") {
  let RootUrl = "https://www.citychain.com.sg/search?options%5Bprefix%5D=last";

  let urls = [];

  for (let i = 1; i <= maxPages; i++) {
    let url = `${RootUrl}&page=${i}&q=${term}`;
    urls.push(url);
  }

  return urls;
}

async function ProductInfo(page, PageUrl) {
  // const browser = await puppeteer.launch({ headless: "new" });
  // const page = await browser.newPage();
  await page.goto(PageUrl);

  let result = {};

  const [e1] = await page.$$("div.product-block.product-block--header > h1");
  const e1Content = await page.evaluate((element) => element.textContent, e1);
  //   console.log(e1Content);
  result.Name = e1Content;
  //
  //
  //
  const [e2] = await page.$$("div > p > span.metafield-multi_line_text_field");
  let e2Content = "";
  if (e2 != undefined) {
    e2Content = await page.evaluate((element) => element.textContent, e2);
  } else {
    const e2 = await page.$$(".product-single__meta .product-block .rte ul li");

    for (let li of e2) {
      let content = await page.evaluate((element) => element.textContent, li);

      e2Content += content + "\n";
    }
  }
  //   console.log(e2Content);
  result.Specs = e2Content;
  //
  //
  //
  const [e3] = await page.$$("div.product__main-photos ");
  const imagesSlider = await e3.$$("div.flickity-slider > div");

  if (imagesSlider.length != 0) {
    let srcs = [];

    for (let imageWraper of imagesSlider) {
      let [img] = await imageWraper.$$("image-element > img");
      let src = await img.getProperty("src");
      src = await src.jsonValue();

      srcs.push(src);
    }

    result.images = [...srcs];
  } else {
    const [imageWraper] = await e3.$$("div.product-main-slide");
    const [img] = await imageWraper.$$("image-element > img");
    let src = await img.getProperty("src");
    src = await src.jsonValue();

    result.images = [src];
  }

  // await browser.close();

  return result;
}

async function PopulateExcelFile(SearchedProducts, indicator) {
  if (!indicator) {
    const workbook = new ExcelJS.Workbook();
    const workSheet = workbook.addWorksheet("Scraped Data");

    workSheet.addRow(["Title", "Name", "Description", "images"]);

    for (let product of SearchedProducts) {
      let { Title, Name, Description, Images } = product;
      // console.log(Image);
      workSheet.addRow([Title, Name, Description, ...Images]);
    }

    await workbook.xlsx.writeFile(SOURCE_FILE);
  } else {
    await updateExcelFile(SearchedProducts, SOURCE_FILE);
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
            let count = 4;
            for (let img of product.Images) {
              row.getCell(count++).value = img;
            }

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
