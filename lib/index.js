const { Workbook } = require("excel4node")
const Sheet = require("./classes/sheet");
const Column = require("./classes/column")

module.exports = { Workbook, Sheet, Column };
module.exports.excelMate = { Workbook, Sheet, Column };