"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.Excel = void 0;
const tslib_1 = require("tslib");
const http_1 = require("./util/http");
const util_1 = require("./util/util");
const xl = require("excel4node");
const node_xlsx_1 = tslib_1.__importDefault(require("node-xlsx"));
const types_1 = require("./types");
const fs = require("fs");
const path = require("path");
class Excel {
    constructor(options = {}) {
        /**
         * file save path
         */
        this.path = path.join(__dirname, "public");
        /**
         * file name
         */
        this.fileName = "excel";
        /**
         * export excel suffix
         */
        this.suffix = "xlsx";
        this.filePath = "";
        /**
         * temporarily file name
         */
        this.fileNameTmp = "";
        /**
         * workSheets
         */
        this.wsList = [];
        this.wsIndex = 0;
        /**
         * row column map
         */
        this.rowColumnMap = {};
        /**
         * current row column item
         */
        this.currentRowColumnItem = {};
        const { debug = false } = options;
        this.http = new http_1.Http();
        this.wb = new xl.Workbook(options);
        this.debug = debug;
    }
    /**
     * add a work sheet
     * @param sheetName
     * @returns
     */
    addWorkSheet(sheetName, options = {}) {
        const initRowColumnItem = {
            row: 1,
            column: 1,
            initCol: 1,
            initRow: 1,
            depthMap: {},
            depth: 1,
        };
        this.ws = this.wb.addWorksheet(sheetName, options);
        this.wsList.push(this.ws);
        this.wsIndex = this.wsList.length - 1;
        this.rowColumnMap[this.wsIndex] = initRowColumnItem;
        this.currentRowColumnItem = initRowColumnItem;
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
     *
     * cell data
     * @param data
     */
    async render(data) {
        for (const cells of data) {
            let maxRowSpan = 0;
            for (const cell of cells) {
                if (!cell["colSpan"])
                    cell["colSpan"] = 1;
                if (!cell["rowSpan"])
                    cell["rowSpan"] = 1;
                this.currentRowColumnItem.initRow = this.currentRowColumnItem.row;
                this.currentRowColumnItem.initCol = this.currentRowColumnItem.column;
                await this.renderCell(cell);
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
    async renderCell(cell) {
        const { childCells = [], colSpan = 1, rowSpan = 1 } = cell;
        this.currentRowColumnItem.depthMap[this.currentRowColumnItem.depth] = {
            row: this.currentRowColumnItem.row,
            column: this.currentRowColumnItem.column,
            colSpan,
            rowSpan,
        };
        if (this.debug)
            console.log(" site ", this.currentRowColumnItem.row, this.currentRowColumnItem.column, cell["type"] || types_1.CellType.string, cell["text"]);
        await this.setCellValue(cell);
        if (childCells.length) {
            this.currentRowColumnItem.row =
                this.currentRowColumnItem.depthMap[this.currentRowColumnItem.depth].row;
            this.currentRowColumnItem.depth++;
            for (const cells of childCells) {
                this.currentRowColumnItem.column =
                    this.currentRowColumnItem.depthMap[this.currentRowColumnItem.depth - 1].column +
                        colSpan -
                        1;
                let maxRowSpan = 0;
                for (const cell of cells) {
                    if (!cell["rowSpan"])
                        cell["rowSpan"] = 1;
                    if (!cell["colSpan"])
                        cell["colSpan"] = 1;
                    this.currentRowColumnItem.column++;
                    await this.renderCell(cell);
                    maxRowSpan = Math.max(maxRowSpan, cell["rowSpan"]);
                }
                if (this.debug)
                    console.log("maxRowSpan", maxRowSpan);
                this.currentRowColumnItem.row += maxRowSpan;
            }
            this.currentRowColumnItem.depth--;
        }
        this.currentRowColumnItem.row =
            this.currentRowColumnItem.depthMap[this.currentRowColumnItem.depth].row;
        this.currentRowColumnItem.column =
            this.currentRowColumnItem.depthMap[this.currentRowColumnItem.depth].column;
        return;
    }
    /**
     * set a cell value
     * @param cell
     */
    async setCellValue(cell) {
        const { text, rowSpan = 1, colSpan = 1, type = types_1.CellType.string, style = {}, } = cell;
        const rowEnd = rowSpan > 1
            ? this.currentRowColumnItem.row + rowSpan - 1
            : this.currentRowColumnItem.row;
        const colEnd = colSpan > 1
            ? this.currentRowColumnItem.column + colSpan - 1
            : this.currentRowColumnItem.column;
        switch (type) {
            case types_1.CellType.string:
                this.ws
                    .cell(this.currentRowColumnItem.row, this.currentRowColumnItem.column, rowEnd, colEnd, true)
                    .string(text)
                    .style(style);
                break;
            case types_1.CellType.image:
                await this.setImage(this.currentRowColumnItem.row, this.currentRowColumnItem.column, rowEnd, colEnd, text, style);
                break;
            case types_1.CellType.number:
                this.ws
                    .cell(this.currentRowColumnItem.row, this.currentRowColumnItem.column, rowEnd, colEnd, true)
                    .number(text)
                    .style(style);
                break;
            case types_1.CellType.date:
                this.ws
                    .cell(this.currentRowColumnItem.row, this.currentRowColumnItem.column, rowEnd, colEnd, true)
                    .date(text)
                    .style(style);
                break;
            case types_1.CellType.link:
                this.ws
                    .cell(this.currentRowColumnItem.row, this.currentRowColumnItem.column, rowEnd, colEnd, true)
                    .link(text)
                    .style(style);
                break;
        }
        if (rowSpan > 1)
            this.currentRowColumnItem.row += rowSpan - 1;
        if (colSpan > 1)
            this.currentRowColumnItem.column += colSpan + 1;
    }
    /**
     * set this file name
     * @param fileName
     * @returns
     */
    setFileName(fileName) {
        this.fileName = fileName;
        return this;
    }
    /**
     * set save path
     * @param path
     * @returns
     */
    setPath(path) {
        this.path = path;
        return this;
    }
    setRowHeight(row, height) {
        this.ws.row(row).setHeight(height);
    }
    setColWidth(col, width) {
        this.ws.row(col).setHeight(width);
    }
    setRowFreeze(rowNumber, autoScrollTo = 0) {
        this.ws.row(rowNumber).freeze(autoScrollTo);
    }
    setColFreeze(colNumber, auToScrollTo = 0) {
        this.ws.column(colNumber).freeze(auToScrollTo);
    }
    setRowHide(row) {
        this.ws.row(row).hide();
    }
    setColHide(col) {
        this.ws.column(col).hide();
    }
    /**
     * save as excel file
     */
    async saveFile() {
        this.fileNameTmp = this.fileName + util_1.createRandomStr(15) + "." + this.suffix;
        this.filePath = path.join(this.path, this.fileNameTmp);
        await this.writeFile();
        return this.filePath;
    }
    /**
     * read a excel and parse to Array
     * @param path
     * @param sheetIndex
     * @returns
     */
    async readExcel(path, sheetIndex = 0) {
        const workSheets = node_xlsx_1.default.parse(path);
        const sheet = workSheets[sheetIndex];
        const data = sheet.data;
        return data;
    }
    /**
     * set cell image
     * @param row
     * @param column
     * @param data
     */
    async setImage(row, column, rowEnd, colEnd, data, style) {
        if (!data)
            this.ws.cell(row, column, rowEnd, colEnd, true).string(data).style(style);
        const res = await this.http.get(data, { responseType: "arraybuffer" });
        const from = {
            row,
            col: column,
            colOff: "0.1in",
            rowOff: 0,
        };
        const to = rowEnd && colEnd
            ? {
                row: rowEnd + 1,
                col: colEnd + 1,
                colOff: "0.1in",
                rowOff: 0,
            }
            : from;
        try {
            this.ws.cell(row, column, rowEnd, colEnd, true).style(style);
            this.ws.addImage({
                image: res.data,
                type: "picture",
                position: {
                    type: "oneCellAnchor",
                    from,
                    to,
                },
            });
        }
        catch (error) {
            throw error;
        }
    }
    writeToBuffer() {
        return this.wb.writeToBuffer();
    }
    /**
     * write to a excel file
     * @param filePath
     * @returns
     */
    async writeFile() {
        return await new Promise((resolver, reject) => {
            this.wb.write(this.filePath, function (error) {
                if (error)
                    reject(error);
                resolver(true);
            });
        });
    }
    /**
     * remove the temporarily excel file
     */
    removeFile() {
        try {
            fs.unlinkSync(this.filePath);
        }
        catch (error) { }
    }
    /**
     * set http context header for export excel
     * @param ctx
     */
    setCtxHeader(ctx) {
        const ua = (ctx.req.headers["user-agent"] || "").toLowerCase();
        let fileName = encodeURIComponent(this.fileName + "." + this.suffix);
        if (ua.indexOf("msie") >= 0 || ua.indexOf("chrome") >= 0) {
            ctx.set("Content-Disposition", `attachment; filename=${fileName}`);
        }
        else if (ua.indexOf("firefox") >= 0) {
            ctx.set("Content-Disposition", `attachment; filename*=${fileName}`);
        }
        else {
            ctx.set("Content-Disposition", `attachment; filename=${Buffer.from(this.fileName + "." + this.suffix).toString("binary")}`);
        }
    }
}
exports.Excel = Excel;
//# sourceMappingURL=excel.js.map