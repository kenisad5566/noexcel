import Excel from "../src/index";
import { Cell } from "../src/types";
const path = require("path");

async function exportFile(): Promise<Buffer> {
  const exportPath = path.join(__dirname, "../tmp");
  const fileName = "test";
  const sheetName = "student";

  const data = [
    [{ text: "序号" }, { text: "姓名" }, { text: "年龄" }, { text: "职位" }],
    [{ text: "1" }, { text: "小明" }, { text: "15" }, { text: "班长" }],
    [{ text: "2" }, { text: "小华" }, { text: "14" }, { text: "学习委员" }],
    [{ text: "3" }, { text: "小爱" }, { text: "14" }, { text: "组长" }],
  ] as Cell[][];

  const excel = new Excel();
  excel.addWorkSheet(sheetName).setSavePath(exportPath).setFileName(fileName);
  await excel.render(data);
  const buffer = await excel.writeToBuffer();

  console.log("buffer", buffer);

  return buffer;
}

exportFile();
//todo: you can return this buffer to ctx for export
