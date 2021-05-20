import { Excel } from "./excel";
import { Cell } from "./types";

const data = [
  [
    {
      text: "https://pic4.zhimg.com/80/v2-e033acbeacba3cd43e4874b1fa34afc8_720w.jpg",
      type: "string",
      rowSpan: 2,
      childCells: [[{ text: "a" }], [{ text: "b" }]],
      style: {
        font: {
          bold: true,
          underline: true,
        },
      },
    },
  ],
] as Cell[][];

async function test() {
  const excel = new Excel({ debug: true });

  await excel
    .addWorkSheet("test")
    .setFileName("tttt")
    .setPath("./tmp")
    .render(data);

  await excel.saveFile();
}

test();
