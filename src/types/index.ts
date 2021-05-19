export interface CellStyle {}

export interface Cell {
  text: string;
  type?: CellType;
  rowSpan?: number;
  colSpan?: number;
  style?: CellStyle;
  childCells?: Cell[][];
}

export enum CellType {
  number = "number",
  string = "string",
  image = "image",
  date = "date",
  link = "link",
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
  depthMap: { [key: string]: Depth };
  depth: number;
}
