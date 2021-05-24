import Excel from "../src/index";
import { Cell } from "../src/types";
const path = require("path");

async function exportFile() {
  const exportPath = path.join(__dirname, "../tmp");
  const fileName = "test";
  const sheetName = "student";

  const data = [
    [{ text: "班级" }, { text: "姓名" }, { text: "年龄" }, { text: "职位" }],
    [
      {
        text: "一年级",
        rowSpan: 3,
        childCells: [
          [{ text: "小明" }, { text: "15" }, { text: "班长" }],
          [{ text: "小周" }, { text: "15" }, { text: "学习委员" }],
          [{ text: "小朋" }, { text: "15" }, { text: "组长" }],
        ],
      },
    ],
    [{ text: "2" }, { text: "小华" }, { text: "14" }, { text: "学习委员" }],
    [{ text: "3" }, { text: "小爱" }, { text: "13" }, { text: "组长" }],
  ] as Cell[][];

  const excel = new Excel();
  excel.addWorkSheet(sheetName).setSavePath(exportPath).setFileName(fileName);
  await excel.render(data);
  await excel.saveFile();
}

exportFile();
