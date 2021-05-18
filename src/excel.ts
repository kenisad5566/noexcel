import { Http } from "./util/http";
import { createRandomStr } from "./util/util";

const xl = require("excel4node");
import xlsx from "node-xlsx";
import { Cell, CellType, RowColumnItem } from "./types";

const fs = require("fs");
const path = require("path");

export class Excel {
  /**
   * file save path
   */
  private path = path.join(__dirname, "public");

  /**
   * file name
   */
  private fileName = "excel";

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
  private ws: any;

  /**
   * workSheets
   */
  private wsList: any[] = [];

  private wsIndex: number = 0;

  /**
   * export excel suffix
   */
  private suffix: string = ".xlsx";

  /**
   * row column map
   */
  private rowColumnMap: { [key: string]: RowColumnItem } = {};

  /**
   * current row column item
   */
  private currentRowColumnItem: RowColumnItem = {} as any;

  /**
   * debug console.log some msg
   */
  private debug;

  constructor(debug: boolean = false) {
    this.http = new Http();
    this.wb = new xl.Workbook();
    this.debug = debug;
  }

  /**
   * add a work sheet
   * @param sheetName
   * @returns
   */
  addWorkSheet(sheetName: string): this {
    const initRowColumnItem: RowColumnItem = {
      row: 1,
      column: 1,
      initCol: 1,
      initRow: 1,
      depthMap: {},
      depth: 1,
    };
    const ws = this.wb.addWorksheet(sheetName);
    this.ws = ws;
    this.wsList.push(ws);
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
   * {type:"text/image/number", text:"内容"}
   * simple excel
   * @param data
   */
  async simpleRender(data: Cell[][]): Promise<this> {
    let row = 1;
    for (const index1 in data) {
      const objects = data[index1];
      for (const index2 in objects) {
        const { text, type } = objects[index2];
        const column = Number(index2 + 1);
        if (type === CellType.number || type === CellType.text)
          this.setText(row, column, text);
        if (type === CellType.image) await this.setImage(text, row, column);
      }
      row++;
    }
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

        this.renderCell(cell);
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

  private async renderCell(cell: any) {
    const { data = [], colSpan, rowSpan } = cell;
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
        cell["text"]
      );
    if (this.debug)
      console.log("this.depth", this.currentRowColumnItem.depthMap);

    this.setCellValue(cell);
    if (data.length) {
      this.currentRowColumnItem.row =
        this.currentRowColumnItem.depthMap[this.currentRowColumnItem.depth].row;
      this.currentRowColumnItem.depth++;

      for (const cells of data) {
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
          this.renderCell(cell);
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
  private setCellValue(cell: Cell) {
    const { text, rowSpan = 1, colSpan = 1, type = CellType.text } = cell;

    const rowEnd =
      rowSpan > 1
        ? this.currentRowColumnItem.row + rowSpan - 1
        : this.currentRowColumnItem.row;
    const colEnd =
      colSpan > 1
        ? this.currentRowColumnItem.column + colSpan - 1
        : this.currentRowColumnItem.column;

    if (type === CellType.text || type === CellType.number)
      this.ws
        .cell(
          this.currentRowColumnItem.row,
          this.currentRowColumnItem.column,
          rowEnd,
          colEnd,
          true
        )
        .string(text.toString());
    if (type === CellType.image)
      this.setImage(
        text,
        this.currentRowColumnItem.row,
        this.currentRowColumnItem.column,
        rowEnd,
        colEnd
      );

    if (rowSpan > 1) this.currentRowColumnItem.row += rowSpan - 1;
    if (colSpan > 1) this.currentRowColumnItem.column += colSpan + 1;
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
  setPath(path: string): this {
    this.path = path;
    return this;
  }

  /**
   * export as buffer, you can handler it with ctx to return to browser
   * @returns
   */
  async exportAsBuffer() {
    await this.saveFile();
    const xlsx = await this.readAsBuffer();
    this.removeFile();

    return xlsx;
  }

  /**
   * save as excel file
   */
  async saveFile() {
    this.fileNameTmp = this.fileName + createRandomStr(15) + this.suffix;
    const filePath = path.join(this.path, this.fileNameTmp);
    await this.writeExcel(filePath);
  }

  /**
   * read as buffer
   * @returns
   */
  async readAsBuffer(): Promise<Buffer> {
    const filePath = path.join(this.path, this.fileNameTmp);
    return await fs.readFileSync(filePath);
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
   * set cell text value
   * @param row
   * @param column
   * @param data
   */
  private setText(row: number, column: number, data: string) {
    this.ws.cell(row, column).string(data);
  }

  /**
   * set cell image
   * @param row
   * @param column
   * @param data
   */
  private async setImage(
    data: string,
    row: number,
    column: number,
    rowEnd: number = 0,
    colEnd: number = 0
  ) {
    if (!data) this.setText(row, column, "no image");
    const res = await this.http.get(data, { responseType: "arraybuffer" });

    fs.writeFileSync("./xxx.png", res.data);

    const from = {
      row,
      col: column,
      colOff: "0.1in",
      rowOff: 0,
    };

    const to =
      rowEnd && colEnd
        ? {
            row: rowEnd,
            col: colEnd,
            colOff: "0.1in",
            rowOff: 0,
          }
        : from;

    console.log("from", from);
    console.log("to", to);

    console.log("res", res.data);

    try {
      this.ws.row(row).setHeight(100);
      console.log("222");

      this.ws.addImage({
        image: res.data,
        type: "picture",
        position: {
          type: "oneCellAnchor",
          from,
          // to,
        },
      });
      console.log("333");
    } catch (error) {
      console.log("error", error);
    }
  }

  /**
   * write to a excel file
   * @param filePath
   * @returns
   */
  private async writeExcel(filePath: string): Promise<boolean> {
    return await new Promise((resolver, reject) => {
      this.wb.write(filePath, function (error: any) {
        if (error) reject(error);
        resolver(true);
      });
    });
  }

  /**
   * remove the temporarily excel file
   */
  private removeFile() {
    const filePath = path.join(this.path, this.fileNameTmp);
    try {
      fs.unlinkSync(filePath);
    } catch (error) {}
  }
}
