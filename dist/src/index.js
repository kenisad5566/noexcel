"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
const excel_1 = require("./excel");
const data = [[{ text: "xxx" }]];
const excel = new excel_1.Excel(true);
excel.addWorkSheet("test").setName("tttt").renderData(data);
//# sourceMappingURL=index.js.map