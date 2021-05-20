import { Excel } from "./excel";
import { Cell } from "./types";

const data = [
  [
    {
      text: "c",
      type: "string",
    },
    {
      text: "c",
      type: "string",
    },
    {
      text: "c",
      type: "string",
    },
    {
      text: "c",
      type: "string",
    },
    {
      text: "c",
      type: "string",
    },
  ],
  [
    {
      text: "a",
      style: {
        font: { color: "red", size: 20 },
      },
    },
    {
      text: "a",
      style: {
        font: { color: "red", size: 20 },
      },
    },
    {
      text: "a",
      style: {
        font: { color: "red", size: 20 },
      },
    },
    {
      text: "a",
      style: {
        font: { color: "red", size: 20 },
      },
    },
  ],
  [{ text: "b" }, { text: "b" }, { text: "b" }, { text: "b" }, { text: "b" }],
] as Cell[][];

async function test() {
  const excel = new Excel({});

  await excel
    .addWorkSheet("test", {})
    .setFileName("tttt")
    .setPath("./tmp")
    .render(data);

  excel.setColHide(1);

  await excel.saveFile();
}

test();
