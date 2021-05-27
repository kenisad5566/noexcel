import Excel from "../src/index";
import { Cell } from "../src/types";
const path = require("path");

async function exportFile(): Promise<Buffer> {
  const exportPath = path.join(__dirname, "../tmp");
  const fileName = "test";
  const sheetName = "student";

  const data = [
    [{ text: "s/n" }, { text: "name" }, { text: "age" }, { text: "position" }],
    [{ text: "1" }, { text: "ming" }, { text: "15" }, { text: "monitor" }],
    [{ text: "2" }, { text: "hua" }, { text: "14" }, { text: "commissary" }],
    [{ text: "3" }, { text: "ai" }, { text: "14" }, { text: "supervisor" }],
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
