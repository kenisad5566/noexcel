import Excel from "../src/index";
import { Cell } from "../src/types";
const path = require("path");

async function exportFile() {
  const exportPath = path.join(__dirname, "../tmp");
  const fileName = "test";
  const sheetName = "student";

  const data = [
    [
      { text: "序号" },
      { text: "姓名" },
      { text: "年龄" },
      { text: "职位" },
      { text: "个人链接" },
    ],
    [
      { text: "1", style: { font: { bold: true } } },
      { text: "小明" },
      { text: "15" },
      { text: "班长" },
      { text: "http://www.google.com", type: "link" },
    ],
    [
      { text: "2" },
      { text: "小华", style: { font: { size: 14 } } },
      { text: "14" },
      { text: "学习委员" },
      { text: "http://www.google.com", type: "string" },
    ],
    [
      { text: "3" },
      { text: "小爱" },
      { text: "13" },
      { text: "组长" },
      {
        text: "http://www.google.com",
        type: "link",
        style: { font: { underline: true, bold: true, color: "black" } },
      },
    ],
  ] as Cell[][];

  const excel = new Excel();
  excel.addWorkSheet(sheetName).setSavePath(exportPath).setFileName(fileName);
  await excel.render(data);
  await excel.saveFile();
}

exportFile();
