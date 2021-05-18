export interface Cell {
    text: string;
    type?: CellType;
    rowSpan?: number;
    colSpan?: number;
    data?: Cell[];
}
export declare enum CellType {
    number = "number",
    text = "text",
    pic = "pic"
}
export interface Depth {
    row: number;
    column: number;
    colSpan: number;
    rowSpan: number;
}
/**
 *  worksheet row and column record map
 */
export interface RowColumnItem {
    row: number;
    column: number;
    initRow: number;
    initCol: number;
    depthMap: {
        [key: string]: Depth;
    };
    depth: number;
}
