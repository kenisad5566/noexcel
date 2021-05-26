/// <reference types="node" />
import { Cell, RowColumnItem, WorkBookOpts, WorkSheetStyle } from "./types";
export declare class Excel {
    /**
     * file save path
     */
    path: any;
    /**
     * file name
     */
    fileName: string;
    /**
     * export excel suffix
     */
    suffix: string;
    private filePath;
    /**
     * temporarily file name
     */
    private fileNameTmp;
    /**
     * http agent
     * use for pic render
     */
    private http;
    /**
     * workbook
     */
    private wb;
    /**
     * workSheet
     */
    ws: any;
    /**
     * workSheets
     */
    wsList: any[];
    wsIndex: number;
    /**
     * row column map
     */
    private rowColumnMap;
    /**
     * current row column item
     */
    currentRowColumnItem: RowColumnItem;
    /**
     * debug console.log some msg
     */
    private debug;
    constructor(options?: WorkBookOpts);
    /**
     * add a work sheet
     * @param sheetName
     * @returns
     */
    addWorkSheet(sheetName: string, options?: WorkSheetStyle): this;
    /**
     * select a work sheet
     * @param index
     * @returns
     */
    selectSheet(index: number): this;
    /**
     *
     * cell data
     * @param data
     */
    render(data: Cell[][]): Promise<this>;
    /**
     * set this file name
     * @param fileName
     * @returns
     */
    setFileName(fileName: string): this;
    /**
     * set save path
     * @param path
     * @returns
     */
    setSavePath(path: string): this;
    setRowHeight(row: number, height: number): void;
    setColWidth(col: number, width: number): void;
    setRowFreeze(rowNumber: number, autoScrollTo?: number): void;
    setColFreeze(colNumber: number, auToScrollTo?: number): void;
    setRowHide(row: number): void;
    setColHide(col: number): void;
    setSuffix(suffix: string): void;
    /**
     * save as excel file
     */
    saveFile(): Promise<string>;
    /**
     * read a excel and parse to Array
     * @param path
     * @param sheetIndex
     * @returns
     */
    readExcel(path: string, sheetIndex?: number): Promise<unknown[][]>;
    /**
     * set cell image
     * @param row
     * @param column
     * @param data
     */
    private setImage;
    writeToBuffer(): Promise<Buffer>;
    /**
     * remove the temporarily excel file
     */
    removeFile(): void;
    /**
     * set http context header for export excel
     * @param ctx
     */
    setCtxHeader(ctx: any): void;
    /**
     * write to a excel file
     * @param filePath
     * @returns
     */
    private writeFile;
    private renderCell;
    /**
     * set a cell value
     * @param cell
     */
    private setCellValue;
}
