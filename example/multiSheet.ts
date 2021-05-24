import Excel from "../src/index";
import { Cell } from "../src/types";
const path = require("path");

async function exportFile() {
  const exportPath = path.join(__dirname, "../tmp");
  const fileName = "test";
  const sheetName1 = "student1";
  const sheetName2 = "student2";

  const data = [
    [{ text: "序号" }, { text: "姓名" }, { text: "年龄" }, { text: "职位" }],
    [{ text: "1" }, { text: "小明" }, { text: "15" }, { text: "班长" }],
    [{ text: "2" }, { text: "小华" }, { text: "14" }, { text: "学习委员" }],
    [{ text: "3" }, { text: "小爱" }, { text: "13" }, { text: "组长" }],
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
