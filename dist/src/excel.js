"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.Excel = void 0;
const tslib_1 = require("tslib");
const http_1 = require("./util/http");
const util_1 = require("./util/util");
const xl = require("excel4node");
const node_xlsx_1 = tslib_1.__importDefault(require("node-xlsx"));
const types_1 = require("../types");
const fs = require("fs");
const path = require("path");
class Excel {
    constructor(debug = false) {
        /**
         * save dir name
         */
        this.dir = "public";
        /**
         * workSheets
         */
        this.wsList = [];
        this.wsIndex = 0;
        /**
         * export excel name
         */
        this.name = "excelName";
        /**
         * export excel suffix
         */
        this.suffix = ".xlsx";
        /**
         * row column map
         */
        this.rowColumnMap = {};
        /**
         * init row column item
         */
        this.initRowColumnItem = {
            row: 1,
            column: 1,
            initCol: 1,
            initRow: 1,
            depthMap: {},
            depth: 1,
        };
        this.currentRowColumnItem = this.initRowColumnItem;
        this.http = new http_1.Http();
        this.wb = new xl.Workbook();
        this.debug = debug;
    }
    /**
     * add a work sheet
     * @param sheetName
     * @returns
     */
    addWorkSheet(sheetName) {
        const ws = this.wb.addWorksheet(sheetName);
        this.ws = ws;
        this.wsList.push(ws);
        this.wsIndex = this.wsList.length - 1;
        this.rowColumnMap[this.wsIndex] = this.initRowColumnItem;
        this.currentRowColumnItem = this.initRowColumnItem;
        return this;
    }
    /**
     * select a work sheet
     * @param index
     * @returns
     */
    selectSheet(index) {
        this.ws = this.wsList[index];
        if (!this.ws)
            throw new Error("no work sheet");
        this.wsIndex = index;
        this.currentRowColumnItem = this.rowColumnMap[this.wsIndex];
        return this;
    }
    /**
     * {type:"text/image", data:"内容"}
     * 简单的表格
     * @param data
     */
    async simpleRender(data) {
        let row = 1;
        for (const index1 in data) {
            const objects = data[index1];
            for (const index2 in objects) {
                const { text, type } = objects[index2];
                const column = Number(index2 + 1);
                if (type === types_1.CellType.number || type === types_1.CellType.text)
                    this.setText(row, column, text);
                if (type === types_1.CellType.pic)
                    await this.setImage(row, column, text);
            }
            row++;
        }
        return this;
    }
    /**
     * {type:"text/image", data:"内容", colSpan:1, rowSpan:1}
     * 简单的表格
     * @param data
     */
    async renderData(data) {
        for (const cells of data) {
            let maxRowSpan = 0;
            for (const cell of cells) {
                if (!cell["colSpan"])
                    cell["colSpan"] = 1;
                if (!cell["rowSpan"])
                    cell["rowSpan"] = 1;
                this.currentRowColumnItem.initRow = this.currentRowColumnItem.row;
                this.currentRowColumnItem.initCol = this.currentRowColumnItem.column;
                this.renderCellVertical(cell);
                this.currentRowColumnItem.column =
                    this.currentRowColumnItem.initCol + cell["colSpan"];
                this.currentRowColumnItem.row = this.currentRowColumnItem.initRow;
                maxRowSpan = Math.max(maxRowSpan, cell["rowSpan"]);
            }
            this.currentRowColumnItem.row += maxRowSpan;
            this.currentRowColumnItem.column = 1;
            this.currentRowColumnItem.depth = 1;
            this.currentRowColumnItem.depthMap = {};
        }
        return this;
    }
    async renderCellVertical(cell) {
        const { data = [], colSpan, rowSpan } = cell;
        this.currentRowColumnItem.depthMap[this.currentRowColumnItem.depth] = {
            row: this.currentRowColumnItem.row,
            column: this.currentRowColumnItem.column,
            colSpan,
            rowSpan,
        };
        if (this.debug)
            console.log(" site ", this.currentRowColumnItem.row, this.currentRowColumnItem.column, cell["text"]);
        if (this.debug)
            console.log("this.depth", this.currentRowColumnItem.depthMap);
        this.setCellValue(cell);
        if (data.length) {
            this.currentRowColumnItem.row =
                this.currentRowColumnItem.depthMap[this.currentRowColumnItem.depth]["row"];
            this.currentRowColumnItem.depth++;
            for (const cells of data) {
                this.currentRowColumnItem.column =
                    this.currentRowColumnItem.depthMap[this.currentRowColumnItem.depth - 1]["column"] +
                        colSpan -
                        1;
                let maxRowSpan = 0;
                for (const cell of cells) {
                    if (!cell["rowSpan"])
                        cell["rowSpan"] = 1;
                    if (!cell["colSpan"])
                        cell["colSpan"] = 1;
                    this.currentRowColumnItem.column++;
                    this.renderCellVertical(cell);
                    maxRowSpan = Math.max(maxRowSpan, cell["rowSpan"]);
                }
                if (this.debug)
                    console.log("maxRowSpan", maxRowSpan);
                this.currentRowColumnItem.row += maxRowSpan;
            }
            this.currentRowColumnItem.depth--;
        }
        this.currentRowColumnItem.row =
            this.currentRowColumnItem.depthMap[this.currentRowColumnItem.depth]["row"];
        this.currentRowColumnItem.column =
            this.currentRowColumnItem.depthMap[this.currentRowColumnItem.depth]["column"];
        return;
    }
    setCellValue(cell) {
        const { text, rowSpan, colSpan } = cell;
        const rowEnd = rowSpan > 1
            ? this.currentRowColumnItem.row + rowSpan - 1
            : this.currentRowColumnItem.row;
        const colEnd = colSpan > 1
            ? this.currentRowColumnItem.column + colSpan - 1
            : this.currentRowColumnItem.column;
        this.ws
            .cell(this.currentRowColumnItem.row, this.currentRowColumnItem.column, rowEnd, colEnd, true)
            .string(text);
        if (rowSpan > 1)
            this.currentRowColumnItem.row += rowSpan - 1;
        if (colSpan > 1)
            this.currentRowColumnItem.column += colSpan + 1;
    }
    setName(name) {
        this.name = name;
        return this;
    }
    async export() {
        const pwd = process.cwd();
        const filePath = path.join(pwd, this.dir, this.name + util_1.createRandomStr(10) + this.suffix);
        await this.writeExcel(filePath);
        const xlsx = await fs.readFileSync(filePath);
        this.removeFile(filePath);
        return xlsx;
    }
    async readExcel(path, sheetIndex = 0) {
        const workSheets = node_xlsx_1.default.parse(path);
        const sheet = workSheets[sheetIndex];
        const data = sheet.data;
        return data;
    }
    setText(row, column, data) {
        this.ws.cell(row, column).string(data);
    }
    async setImage(row, column, data) {
        if (!data)
            this.setText(row, column, "没有图片");
        const res = await this.http.get(data, { responseType: "arraybuffer" });
        try {
            this.ws.row(row).setHeight(100);
            this.ws.addImage({
                image: res.data,
                type: "picture",
                position: {
                    type: "oneCellAnchor",
                    from: {
                        row,
                        col: column,
                        colOff: "0.1in",
                        rowOff: 0,
                    },
                },
            });
        }
        catch (error) { }
    }
    async writeExcel(filePath) {
        return await new Promise((resolver, reject) => {
            this.wb.write(filePath, function (error) {
                if (error)
                    reject(error);
                resolver(true);
            });
        });
    }
    removeFile(filePath) {
        fs.unlink(filePath, function (error) {
            if (error)
                console.log("error", error);
        });
    }
}
exports.Excel = Excel;
//# sourceMappingURL=excel.js.map