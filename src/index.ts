import { Excel } from "./excel";
import { Cell } from "./types";

// const data = [
//   [
//     {
//       text: "span big",
//       rowSpan: 3,
//       colSpan: 3,
//       data: [[{ text: "a" }], [{ text: "b" }], [{ text: "c" }]],
//     },
//   ],
//   [{ text: "d" }, { text: "e" }, { text: "f" }],
//   [
//     {
//       text: "https://gimg2.baidu.com/image_search/src=http%3A%2F%2Fi0.hdslb.com%2Fbfs%2Farticle%2F878a6c57bed136d9d176a6eb8289a04787b126bf.jpg&refer=http%3A%2F%2Fi0.hdslb.com&app=2002&size=f9999,10000&q=a80&n=0&g=0n&fmt=jpeg?sec=1623920595&t=53239535f77516e2cd8de119fb95947d",
//       type: "image",
//     },
//   ],
// ] as Cell[][];

const data = [
  [
    {
      text: "https://pic4.zhimg.com/80/v2-e033acbeacba3cd43e4874b1fa34afc8_720w.jpg",
      type: "image",
    },
  ],
] as Cell[][];

async function test() {
  const excel = new Excel(true);

  await excel
    .addWorkSheet("test")
    .setFileName("tttt")
    .setPath(".")
    .render(data);

  await excel.saveFile();
}

test();
