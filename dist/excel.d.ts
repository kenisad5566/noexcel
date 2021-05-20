/// <reference types="node" />
import { Cell, WorkBookOpts, WorkSheetStyle } from "./types";
export declare class Excel {
    /**
     * file save path
     */
    private path;
    /**
     * file name
     */
    private fileName;
    /**
     * export excel suffix
     */
    private suffix;
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
    private ws;
    /**
     * workSheets
     */
    private wsList;
    private wsIndex;
    /**
     * row column map
     */
    private rowColumnMap;
    /**
     * current row column item
     */
    private currentRowColumnItem;
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
    private renderCell;
    /**
     * set a cell value
     * @param cell
     */
    private setCellValue;
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
    setPath(path: string): this;
    setRowHeight(row: number, height: number): void;
    setColWidth(col: number, width: number): void;
    setRowFreeze(rowNumber: number, autoScrollTo?: number): void;
    setColFreeze(colNumber: number, auToScrollTo?: number): void;
    setRowHide(row: number): void;
    setColHide(col: number): void;
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
    writeToBuffer(): Buffer;
    /**
     * write to a excel file
     * @param filePath
     * @returns
     */
    private writeFile;
    /**
     * remove the temporarily excel file
     */
    removeFile(): void;
    /**
     * set http context header for export excel
     * @param ctx
     */
    setCtxHeader(ctx: any): void;
}
