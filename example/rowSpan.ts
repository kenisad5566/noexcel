import Excel from "../src/index";
import { Cell } from "../src/types";
const path = require("path");

async function exportFile() {
  const exportPath = path.join(__dirname, "../tmp");
  const fileName = "test";
  const sheetName = "student";

  const data = [
    [
      { text: "class" },
      { text: "name" },
      { text: "age" },
      { text: "position" },
    ],
    [
      {
        text: "class one",
        rowSpan: 3,
        childCells: [
          [{ text: "ming" }, { text: "15" }, { text: "monitor" }],
          [{ text: "ai" }, { text: "15" }, { text: "commissary" }],
          [{ text: "ai" }, { text: "15" }, { text: "supervisor" }],
        ],
      },
    ],
    [{ text: "2" }, { text: "hua" }, { text: "14" }, { text: "commissary" }],
    [{ text: "3" }, { text: "ai" }, { text: "13" }, { text: "supervisor" }],
  ] as Cell[][];

  const excel = new Excel();
  excel.addWorkSheet(sheetName).setSavePath(exportPath).setFileName(fileName);
  await excel.render(data);
  await excel.saveFile();
}

exportFile();
