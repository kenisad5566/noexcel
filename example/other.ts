import Excel from "../src/index";
import { Cell } from "../src/types";
const path = require("path");

async function exportFile() {
  const exportPath = path.join(__dirname, "../tmp");
  const fileName = "test";
  const sheetName = "student";

  const data = [
    [
      { text: "s/n" },
      { text: "name" },
      { text: "age" },
      { text: "position" },
      { text: "date" },
      { text: "link" },
    ],
    [
      { text: "1", style: { font: { bold: true } } },
      { text: "ming" },
      { text: 15, type: "number" },
      { text: "monitor" },
      { text: new Date(), type: "date", style: { numberFormat: "yyyy-mm-dd" } },
      { text: "http://www.google.com", type: "link" },
    ],
    [
      { text: "2" },
      { text: "hua", style: { font: { size: 14 } } },
      { text: 14, type: "number" },
      { text: "commissary" },
      { text: new Date(), type: "date", style: { numberFormat: "yyyy-mm-dd" } },
      { text: "http://www.google.com", type: "string" },
    ],
    [
      { text: "3" },
      { text: "ai" },
      { text: 13 },
      { text: "supervisor" },
      { text: new Date(), type: "date", style: { numberFormat: "yyyy-mm-dd" } },
      {
        text: "http://www.google.com",
        type: "link",
        style: { font: { underline: true, bold: true, color: "black" } },
      },
    ],
    [
      { text: "4" },
      { text: "ai" },
      { text: 14 },
      { text: "supervisor" },
      { text: new Date(), type: "date", style: { numberFormat: "yyyy-mm-dd" } },
      {
        text: "http://www.cnlogo8.com/d/file/2021-05-20/97517b732413c71921c3a55726f4f299.png",
        type: "image",
      },
    ],
  ] as Cell[][];

  const excel = new Excel();
  excel.addWorkSheet(sheetName).setSavePath(exportPath).setFileName(fileName);
  await excel.render(data);
  excel.setRowHeight(5, 300).setColWidth(6, 70);

  await excel.saveFile();
}

exportFile();
