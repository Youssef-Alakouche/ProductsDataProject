const puppeteer = require("puppeteer");
const { productsFun } = require("./Products");
const { NotFoundedProductsFun } = require("./NotFound");
const ExcelJS = require("exceljs");

const NOT_FOUND = "not found";

let ScrapedDataIndicator = false;

// const products = [
//   "NY0138-14X",
//   "NY0099-81XB",
//   "NH8350-83AB",
//   "NB6021-68L",
//   "FE7040-53E",
//   "EU6094-53A",
//   "ER0216-59D",
//   "ER0214-54D",
// ];

const SOURCE_FILE = "./Search Products/Data.xlsx";
const NOT_FOUND_FILE = "./Search products/NotFound.xlsx";

const SearchedProductsArray = [];

// main();

async function getExcelFileFromH2HubWatches() {
  const browser = await puppeteer.launch({ headless: "new" });
  const page = await browser.newPage();

  //   const products = ["BI5095-05E"];
  //   console.log("start");

  let products = await productsFun(SOURCE_FILE);
  products.reverse();

  //   console.log(products);
  let counter = 0;

  for (let { Title: productTitle, IsScraped } of products) {
    if (!IsScraped) {
      console.log(++counter);
      const url = urls(productTitle);

      //   console.log(productTitle);

      const productPage = await getProductPage(page, url);

      if (productPage == NOT_FOUND) {
        //   console.log(NOT_FOUND);
        SearchedProductsArray.push({
          Title: productTitle,
          Name: NOT_FOUND,
          Description: NOT_FOUND,
          Images: [NOT_FOUND],
        });
      } else {
        const productInfo = await ProductInfo(page, productTitle, productPage);

        if (productInfo == NOT_FOUND) {
          SearchedProductsArray.push({
            Title: productTitle,
            Name: NOT_FOUND,
            Description: NOT_FOUND,
            Images: [NOT_FOUND],
          });
        } else {
          let { Name, Description, Images } = productInfo;

          // console.log(Name);
          // console.log(Images);

          SearchedProductsArray.push({
            Title: productTitle,
            Name,
            Description,
            Images: [...Images],
          });
        }
      }
    } else {
      ScrapedDataIndicator = true;
    }

    //   console.log(productPage);
  }

  //   console.log(SearchedProductsArray);

  await PopulateExcelFile(SearchedProductsArray, ScrapedDataIndicator);
  //   console.log("finish");

  await browser.close();
}

async function getProductPage(page, pageUrl) {
  await page.goto(pageUrl);

  let [e] = await page.$$(
    "#shopify-section-search_page > div > div.row > div > div.products.nt_products_holder.row.fl_center.row_pr_1.cdt_des_1.round_cd_false.nt_cover.ratio_nt.position_8.space_30.nt_default > div:nth-child(1) > div > div.product-image.pr.oh.lazyloaded > div.pr.oh > a"
  );

  if (e == undefined) {
    return NOT_FOUND;
  }

  let href = await e.getProperty("href");
  href = await href.jsonValue();

  return href;
}

async function ProductInfo(page, productTitle, PageUrl) {
  await page.goto(PageUrl);

  let result = {};

  let [RightTitle] = await page.$x('//*[@id="pr_sku_ppr"]');
  RightTitle = await page.evaluate((e) => e.textContent, RightTitle);

  if (RightTitle.trim().toLowerCase() != productTitle.trim().toLowerCase()) {
    // console.log(NOT_FOUND);
    return NOT_FOUND;
  }
  //   ---------------------------------

  let [Name] = await page.$x('//*[@id="shopify-section-pr_summary"]/h1');
  Name = await page.evaluate((e) => e.textContent, Name);

  result.Name = Name;
  // ----------------------------------------------
  try {
    await page.click(
      "div.col-md-6.col-12.pr.product-images.img_action_zoom.pr_sticky_img > div > div.col-12.col-lg.col_thumb > div.p_group_btns.pa.flex > button"
    );
  } catch {
    await page.click(
      "div.col-md-6.col-12.pr.product-images.img_action_zoom.pr_sticky_img > div > div.col-12 > div.p_group_btns.pa.flex > button"
    );
  }

  let counter = 0;
  let Images = [];

  do {
    await page.waitForTimeout(400);
    counter += 1;
    Images = await page.evaluate(() => {
      const imgElements = document.querySelectorAll(".pswp__img");

      // Replace with your appropriate selector
      const srcs = Array.from(imgElements)
        .map((img) => img.getAttribute("src"))
        .filter((img) => img != null);
      return srcs;
    });
  } while (Images.length == 0 && counter < 5);

  //   console.log(Images);
  //   console.log(Images);
  result.Images = [...Images];

  //   ---------------------------------------------

  const [e] = await page.$x('//*[@id="tab_pr_deskl"]/div[3]');
  const [table] = await e.$$("table");
  const [tbody] = await table.$$("tbody");
  const trs = await tbody.$$("tr");

  let text = "";
  for (let tr of trs) {
    let tds = await tr.$$("td");
    for (let i = 0; i < tds.length; i++) {
      let td = tds[i];
      text += await page.evaluate((e) => e.textContent, td);
      text = text.trim();
      if (i != tds.length - 1) {
        text += " : ";
      }
    }

    text += "\n";
  }

  result.Description = text;

  return { ...result };
}

function urls(term) {
  let RootUrl =
    "https://www.h2hubwatches.com/search?type=product&options%5Bunavailable_products%5D=hide&options%5Bprefix%5D=none&q=";

  let url = `${RootUrl}${term}`;

  return url;
}

async function PopulateExcelFile(SearchedProducts, indicator) {
  if (!indicator) {
    const workbook = new ExcelJS.Workbook();
    const workSheet = workbook.addWorksheet("Scraped Data");

    workSheet.addRow(["Title", "Name", "Description", "images"]);

    for (let product of SearchedProducts) {
      let { Title, Name, Description, Images } = product;

      workSheet.addRow([Title, Name, Description, ...Images]);

      // console.log(Image);
    }

    await workbook.xlsx.writeFile(SOURCE_FILE);
  } else {
    await updateExcelFile(SearchedProducts, SOURCE_FILE);
  }

  // await NotFoundProductExcelFile(NOT_FOUND_FILE);

  //   await EliminateNotFoundInData(SearchedProducts);
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

// async function EliminateNotFoundInData(SearchedProducts) {
//   const workbook = new ExcelJS.Workbook();
//   const workSheet = workbook.addWorksheet("Scraped Data");

//   console.log(SearchedProducts);
//   workSheet.addRow(["Title", "Name", "Description", "images"]);

//   let products = await productsFun(NOT_FOUND_FILE);
//   console.log(products);
//   products.reverse();

//   for (let product of SearchedProducts) {
//     let { Title, Name, Description, Images } = product;
//     let title = Title;
//     let Test = false;
//     for (let { Title: NotFoundProductTitle } of products) {
//       if (title == NotFoundProductTitle) {
//         Test = true;
//         break;
//       }
//     }

//     if (!Test) {
//       workSheet.addRow([Title, Name, Description, ...Images]);
//     }
//   }

//   await workbook.xlsx.writeFile(SOURCE_FILE);
// }

module.exports = { getExcelFileFromH2HubWatches };
