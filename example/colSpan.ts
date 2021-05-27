import Excel from "../src/index";
import { Cell } from "../src/types";
const path = require("path");

async function exportFile() {
  const exportPath = path.join(__dirname, "../tmp");
  const fileName = "test";
  const sheetName = "student";

  const data = [
    [
      { text: "class", colSpan: 2 },
      { text: "name" },
      { text: "age" },
      { text: "position" },
    ],
    [
      {
        text: "class 1",
        colSpan: 2,
      },
      { text: "hua" },
      { text: "14" },
      { text: "commissary" },
    ],
    [
      { text: "class 2", colSpan: 2 },
      { text: "hua" },
      { text: "14" },
      { text: "commissary" },
    ],
    [
      {
        text: "class 3",
        colSpan: 2,
        rowSpan: 2,
        childCells: [
          [{ text: "ai" }, { text: "13" }, { text: "supervisor" }],
          [{ text: "ai" }, { text: "13" }, { text: "supervisor" }],
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
