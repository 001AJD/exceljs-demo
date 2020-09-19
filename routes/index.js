var express = require("express");
var router = express.Router();
const Excel = require("exceljs");

/* GET home page. */
router.get("/", function (req, res, next) {
  res.render("index", { title: "Express" });
});

/* GET home page. */
router.get("/get-excel", async function (req, res, next) {
  const workbook = new Excel.Workbook();
  const worksheet = workbook.addWorksheet("sheet", {
    pageSetup: { paperSize: 9, orientation: "landscape" },
  });
  worksheet.columns = [
    { header: "Id", key: "id", width: 10 },
    { header: "Name", key: "name", width: 32 },
    { header: "D.O.B.", key: "dob", width: 15 },
  ];

  // add your rows here
  worksheet.addRow({ id: 1, name: "John Doe", dob: new Date(1970, 1, 1) });
  worksheet.addRow({ id: 2, name: "Jane Doe", dob: new Date(1965, 1, 7) });

  // excel file name
  const fileName = "./excels/myReport" + new Date().getTime() + ".xlsx";

  // save under export.xlsx
  await workbook.xlsx.writeFile(fileName);

  res.download(fileName);
});

module.exports = router;
