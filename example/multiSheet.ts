import Excel from "../src/index";
import { Cell } from "../src/types";
const path = require("path");

async function exportFile() {
  const exportPath = path.join(__dirname, "../tmp");
  const fileName = "test";
  const sheetName1 = "student1";
  const sheetName2 = "student2";

  const data = [
    [{ text: "s/n" }, { text: "name" }, { text: "age" }, { text: "position" }],
    [{ text: "1" }, { text: "ming" }, { text: "15" }, { text: "monitor" }],
    [{ text: "2" }, { text: "hua" }, { text: "14" }, { text: "commissary" }],
    [{ text: "3" }, { text: "ai" }, { text: "13" }, { text: "supervisor" }],
  ] as Cell[][];

  const excel = new Excel();
  excel.setSavePath(exportPath).setFileName(fileName);

  excel.addWorkSheet(sheetName1);
  await excel.render(data);

  excel.addWorkSheet(sheetName2);
  await excel.render(data);

  await excel.saveFile();
}

exportFile();
