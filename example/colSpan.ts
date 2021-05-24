import Excel from "../src/index";
import { Cell } from "../src/types";
const path = require("path");

async function exportFile() {
  const exportPath = path.join(__dirname, "../tmp");
  const fileName = "test";
  const sheetName = "student";

  const data = [
    [
      { text: "班级", colSpan: 2 },
      { text: "姓名" },
      { text: "年龄" },
      { text: "职位" },
    ],
    [
      {
        text: "一年级",
        colSpan: 2,
      },
      { text: "小华" },
      { text: "14" },
      { text: "学习委员" },
    ],
    [
      { text: "二年级", colSpan: 2 },
      { text: "小华" },
      { text: "14" },
      { text: "学习委员" },
    ],
    [
      {
        text: "三年级",
        colSpan: 2,
        rowSpan: 2,
        childCells: [
          [{ text: "小爱" }, { text: "13" }, { text: "组长" }],
          [{ text: "小黄" }, { text: "13" }, { text: "组长" }],
        ],
      },
    ],
  ] as Cell[][];

  const excel = new Excel();
  excel.addWorkSheet(sheetName).setSavePath(exportPath).setFileName(fileName);
  await excel.render(data);
  await excel.saveFile();
}

exportFile();
