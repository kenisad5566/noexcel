"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const excel_1 = require("./excel");
const data = [[{ text: "xxx" }]];
async function test() {
    const excel = new excel_1.Excel(true);
    await (await excel.addWorkSheet("test").setName("tttt").renderData(data)).export();
}
test();
//# sourceMappingURL=index.js.map