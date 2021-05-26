import { Http } from "./util/http";
import { createRandomStr } from "./util/util";
import xlsx from "node-xlsx";
import {
  Cell,
  CellType,
  RowColumnItem,
  WorkBookOpts,
  WorkSheetStyle,
} from "./types";

const xl = require("excel4node");
const fs = require("fs");
const path = require("path");

export class Excel {
  /**
   * file save path
   */
  public path = path.join(__dirname, "public");

  /**
   * file name
   */
  public fileName = "excel";

  /**
   * export excel suffix
   */
  public suffix: string = "xlsx";

  private filePath: string = "";

  /**
   * temporarily file name
   */
  private fileNameTmp = "";

  /**
   * http agent
   * use for pic render
   */
  private http: Http;

  /**
   * workbook
   */
  private wb: any;

  /**
   * workSheet
   */
  public ws: any;

  /**
   * workSheets
   */
  public wsList: any[] = [];

  public wsIndex: number = 0;

  /**
   * row column map
   */
  private rowColumnMap: { [key: string]: RowColumnItem } = {};

  /**
   * current row column item
   */
  public currentRowColumnItem: RowColumnItem = {} as any;

  /**
   * debug console.log some msg
   */
  private debug;

  constructor(options: WorkBookOpts = {}) {
    const { debug = false } = options;
    this.http = new Http();
    this.wb = new xl.Workbook(options);
    this.debug = debug;
  }

  /**
   * add a work sheet
   * @param sheetName
   * @returns
   */
  addWorkSheet(sheetName: string, options: WorkSheetStyle = {}): this {
    const initRowColumnItem: RowColumnItem = {
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
  selectSheet(index: number): this {
    this.ws = this.wsList[index];
    if (!this.ws) throw new Error("no work sheet");
    this.wsIndex = index;
    this.currentRowColumnItem = this.rowColumnMap[this.wsIndex];
    return this;
  }

  /**
   *
   * cell data
   * @param data
   */
  async render(data: Cell[][]): Promise<this> {
    for (const cells of data) {
      let maxRowSpan = 0;
      for (const cell of cells) {
        if (!cell["colSpan"]) cell["colSpan"] = 1;
        if (!cell["rowSpan"]) cell["rowSpan"] = 1;
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

  /**
   * set this file name
   * @param fileName
   * @returns
   */
  setFileName(fileName: string): this {
    this.fileName = fileName;
    return this;
  }

  /**
   * set save path
   * @param path
   * @returns
   */
  setSavePath(path: string): this {
    this.path = path;
    return this;
  }

  setRowHeight(row: number, height: number): this {
    this.ws.row(row).setHeight(height);
    return this;
  }

  setColWidth(col: number, width: number): this {
    this.ws.column(col).setWidth(width);
    return this;
  }

  setRowFreeze(rowNumber: number, autoScrollTo: number = 0): this {
    this.ws.row(rowNumber).freeze(autoScrollTo);
    return this;
  }

  setColFreeze(colNumber: number, auToScrollTo: number = 0): this {
    this.ws.column(colNumber).freeze(auToScrollTo);
    return this;
  }

  setRowHide(row: number): this {
    this.ws.row(row).hide();
    return this;
  }

  setColHide(col: number): this {
    this.ws.column(col).hide();
    return this;
  }

  setSuffix(suffix: string): this {
    this.suffix = suffix;
    return this;
  }

  /**
   * save as excel file
   */
  async saveFile(): Promise<string> {
    this.fileNameTmp = this.fileName + createRandomStr(15) + "." + this.suffix;
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
  async readExcel(path: string, sheetIndex = 0) {
    const workSheets = xlsx.parse(path);
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
  private async setImage(
    row: number,
    column: number,
    rowEnd: number,
    colEnd: number,
    data: string,
    style: any
  ) {
    if (!data)
      this.ws.cell(row, column, rowEnd, colEnd, true).string(data).style(style);
    const res = await this.http.get(data, { responseType: "arraybuffer" });

    const from = {
      row,
      col: column,
      colOff: "0.1in",
      rowOff: 0,
    };

    const to =
      rowEnd && colEnd
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
    } catch (error) {
      throw error;
    }
  }

  public async writeToBuffer(): Promise<Buffer> {
    return await this.wb.writeToBuffer();
  }

  /**
   * remove the temporarily excel file
   */
  public removeFile() {
    try {
      fs.unlinkSync(this.filePath);
    } catch (error) {}
  }

  /**
   * set http context header for export excel
   * @param ctx
   */
  setCtxHeader(ctx: any) {
    const ua = (ctx.req.headers["user-agent"] || "").toLowerCase();
    let fileName = encodeURIComponent(this.fileName + "." + this.suffix);
    if (ua.indexOf("msie") >= 0 || ua.indexOf("chrome") >= 0) {
      ctx.set("Content-Disposition", `attachment; filename=${fileName}`);
    } else if (ua.indexOf("firefox") >= 0) {
      ctx.set("Content-Disposition", `attachment; filename*=${fileName}`);
    } else {
      ctx.set(
        "Content-Disposition",
        `attachment; filename=${Buffer.from(
          this.fileName + "." + this.suffix
        ).toString("binary")}`
      );
    }
  }

  /**
   * write to a excel file
   * @param filePath
   * @returns
   */
  private async writeFile(): Promise<boolean> {
    return await new Promise((resolver, reject) => {
      this.wb.write(this.filePath, function (error: any) {
        if (error) reject(error);
        resolver(true);
      });
    });
  }

  private async renderCell(cell: Cell) {
    const { childCells = [], colSpan = 1, rowSpan = 1 } = cell;
    this.currentRowColumnItem.depthMap[this.currentRowColumnItem.depth] = {
      row: this.currentRowColumnItem.row,
      column: this.currentRowColumnItem.column,
      colSpan,
      rowSpan,
    };

    if (this.debug)
      console.log(
        " site ",
        this.currentRowColumnItem.row,
        this.currentRowColumnItem.column,
        cell["type"] || CellType.string,
        cell["text"]
      );

    await this.setCellValue(cell);
    if (childCells.length) {
      this.currentRowColumnItem.row =
        this.currentRowColumnItem.depthMap[this.currentRowColumnItem.depth].row;
      this.currentRowColumnItem.depth++;

      for (const cells of childCells) {
        this.currentRowColumnItem.column =
          this.currentRowColumnItem.depthMap[
            this.currentRowColumnItem.depth - 1
          ].column +
          colSpan -
          1;
        let maxRowSpan = 0;
        for (const cell of cells) {
          if (!cell["rowSpan"]) cell["rowSpan"] = 1;
          if (!cell["colSpan"]) cell["colSpan"] = 1;
          this.currentRowColumnItem.column++;
          await this.renderCell(cell);
          maxRowSpan = Math.max(maxRowSpan, cell["rowSpan"]);
        }
        if (this.debug) console.log("maxRowSpan", maxRowSpan);

        this.currentRowColumnItem.row += maxRowSpan;
      }
      this.currentRowColumnItem.depth--;
    }

    this.currentRowColumnItem.row =
      this.currentRowColumnItem.depthMap[this.currentRowColumnItem.depth].row;
    this.currentRowColumnItem.column =
      this.currentRowColumnItem.depthMap[
        this.currentRowColumnItem.depth
      ].column;
    return;
  }

  /**
   * set a cell value
   * @param cell
   */
  private async setCellValue(cell: Cell) {
    let {
      text,
      rowSpan = 1,
      colSpan = 1,
      type = CellType.string,
      style = {},
    } = cell;

    const rowEnd =
      rowSpan > 1
        ? this.currentRowColumnItem.row + rowSpan - 1
        : this.currentRowColumnItem.row;
    const colEnd =
      colSpan > 1
        ? this.currentRowColumnItem.column + colSpan - 1
        : this.currentRowColumnItem.column;

    if (typeof text === "number") type = CellType.number;

    switch (type) {
      case CellType.string:
        this.ws
          .cell(
            this.currentRowColumnItem.row,
            this.currentRowColumnItem.column,
            rowEnd,
            colEnd,
            true
          )
          .string(text)
          .style(style);
        break;

      case CellType.image:
        await this.setImage(
          this.currentRowColumnItem.row,
          this.currentRowColumnItem.column,
          rowEnd,
          colEnd,
          text,
          style
        );
        break;

      case CellType.number:
        this.ws
          .cell(
            this.currentRowColumnItem.row,
            this.currentRowColumnItem.column,
            rowEnd,
            colEnd,
            true
          )
          .number(text)
          .style(style);
        break;

      case CellType.date:
        this.ws
          .cell(
            this.currentRowColumnItem.row,
            this.currentRowColumnItem.column,
            rowEnd,
            colEnd,
            true
          )
          .date(text)
          .style(style);
        break;

      case CellType.link:
        this.ws
          .cell(
            this.currentRowColumnItem.row,
            this.currentRowColumnItem.column,
            rowEnd,
            colEnd,
            true
          )
          .link(text)
          .style(style);
        break;
    }

    if (rowSpan > 1) this.currentRowColumnItem.row += rowSpan - 1;
    if (colSpan > 1) this.currentRowColumnItem.column += colSpan + 1;
  }
}
