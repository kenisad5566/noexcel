import { Cell } from "./types";
export declare class Excel {
    /**
     * save dir name
     */
    private dir;
    /**
     * http agent
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
     * export excel name
     */
    private name;
    /**
     * export excel suffix
     */
    private suffix;
    /**
     * row column map
     */
    private rowColumnMap;
    /**
     * init row column item
     */
    private initRowColumnItem;
    private currentRowColumnItem;
    /**
     * debug console.log some msg
     */
    private debug;
    constructor(debug?: boolean);
    /**
     * add a work sheet
     * @param sheetName
     * @returns
     */
    addWorkSheet(sheetName: string): this;
    /**
     * select a work sheet
     * @param index
     * @returns
     */
    selectSheet(index: number): this;
    /**
     * {type:"text/image", data:"内容"}
     * 简单的表格
     * @param data
     */
    simpleRender(data: Cell[][]): Promise<this>;
    /**
     * {type:"text/image", data:"内容", colSpan:1, rowSpan:1}
     * 简单的表格
     * @param data
     */
    renderData(data: any[][]): Promise<this>;
    private renderCellVertical;
    private setCellValue;
    setName(name: string): this;
    export(): Promise<any>;
    readExcel(path: string, sheetIndex?: number): Promise<unknown[][]>;
    private setText;
    private setImage;
    private writeExcel;
    private removeFile;
}
