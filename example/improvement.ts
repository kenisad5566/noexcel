import { NoExcel, Cell } from "../src/index";
const path = require("path");

const n = 200;
const m = 20;
const data: Cell[][] = [];

const cell = {
  text: "sdfafawerfaweeawerawerwerwerw啥打法鸡撒代付款佳恩阿斯蒂芬",
  rowSpan: 3,
  colSpan: 2,
};

for (let i = 0; i < n; i++) {
  const cells = [];
  for (let j = 0; j < m; j++) {
    cells.push(cell);
  }
  data.push(cells);
}

async function exportFile() {
  const start = new Date().getTime();
  const exportPath = path.join(__dirname, "../tmp");
  const fileName = "test";
  const sheetName = "student";

  const excel = new NoExcel();
  const time2 = new Date().getTime();
  excel.addWorkSheet(sheetName).setSavePath(exportPath).setFileName(fileName);
  const time3 = new Date().getTime();

  await excel.render(data);
  const time4 = new Date().getTime();

  await excel.saveFile();
  const time5 = new Date().getTime();

  const end = new Date().getTime();
  console.log(
    n,
    "use",
    `new object use ${time3 - time2}`,
    `render use ${time4 - time3}`,
    `saveFile use ${time5 - time4}`,
    end - start
  );
}

exportFile();
