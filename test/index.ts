import Excel from "./../dist";
import { Cell } from "./../dist/types";

console.log("ddd", Excel);

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
  console.log("dd", Excel);

  const excel = new Excel({ debug: true });

  await excel
    .addWorkSheet("test", {})
    .setFileName("tttt")
    .setPath("../tmp")
    .render(data);

  excel.setColHide(1);

  await excel.saveFile();
}

test();
