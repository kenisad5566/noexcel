import { Excel } from "./excel";
import { Cell } from "./types";

const data = [
  [
    {
      text: "https://pic4.zhimg.com/80/v2-e033acbeacba3cd43e4874b1fa34afc8_720w.jpg",
      type: "image",
      rowSpan: 2,
      childCells: [[{ text: "a" }], [{ text: "b" }]],
    },
  ],
] as Cell[][];

async function test() {
  const excel = new Excel(true);

  await excel
    .addWorkSheet("test", {
      font: {
        bold: true,
        underline: true,
      },
    })
    .setFileName("tttt")
    .setPath("./tmp")
    .render(data);

  await excel.saveFile();
}

test();
