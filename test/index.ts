import { Excel } from "../src/excel";
import { Cell } from "";

const data = [[{ text: "xxx" }]] as Cell[][];

async function a() {
  const excel = new Excel(true);

  (await excel.addWorkSheet("test").setName("tttt").renderData(data)).export();
}

a();
