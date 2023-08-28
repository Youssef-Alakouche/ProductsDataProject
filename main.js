const { getExcelFileFromCityChain } = require("./citychain");
const { getExcelFileFromThongsia } = require("./thongsia");
const { getExcelFileFromH2HubWatches } = require("./h2hubwatches");

(async () => {
  try {
    console.log("City Chain : Start");
    await getExcelFileFromCityChain();
    console.log("City Chain: finish");
  } catch {
    console.log("city chain");
  }

  console.log("\n66666666666666666666666666666666666666666\n");

  try {
    console.log("Thongsia : Start");
    await getExcelFileFromThongsia();
    console.log("Thongsia : finish");
  } catch {
    console.log("thongsia");
  }

  console.log("\n66666666666666666666666666666666666666666\n");

  try {
    console.log("H2 Hub Watches : Start");
    await getExcelFileFromH2HubWatches();
    console.log("H2 Hub Watches : finish");
  } catch {
    console.log("H2Hubwatches");
  }

  return true;
})();
