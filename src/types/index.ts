export interface CellStyle {
  alignment?: {
    // §18.8.1
    horizontal?:
      | "center"
      | "centerContinuous"
      | "distributed"
      | "fill"
      | "general"
      | "justify"
      | "left"
      | "right";
    indent?: number; // Number of spaces to indent = indent value * 3
    justifyLastLine?: boolean;
    readingOrder?: "contextDependent" | "leftToRight" | "rightToLeft";
    relativeIndent?: number; // number of additional spaces to indent
    shrinkToFit?: boolean;
    textRotation?: number; // number of degrees to rotate text counter-clockwise
    vertical?: "bottom" | "center" | "distributed" | "justify" | "top";
    wrapText?: boolean;
  };
  font?: {
    // §18.8.22
    bold?: boolean;
    charset?: number;
    color?: string;
    condense?: boolean;
    extend?: boolean;
    family?: string;
    italics?: boolean;
    name?: string;
    outline?: boolean;
    scheme?: string; // §18.18.33 ST_FontScheme (Font scheme Styles)
    shadow?: boolean;
    strike?: boolean;
    size?: number;
    underline?: boolean;
    vertAlign?: string; // §22.9.2.17 ST_VerticalAlignRun (Vertical Positioning Location)
  };
  border?: {
    // §18.8.4 border (Border)
    left?: {
      style?: string; //§18.18.3 ST_BorderStyle (Border Line Styles) ['none', 'thin', 'medium', 'dashed', 'dotted', 'thick', 'double', 'hair', 'mediumDashed', 'dashDot', 'mediumDashDot', 'dashDotDot', 'mediumDashDotDot', 'slantDashDot']
      color?: string; // HTML style hex value
    };
    right?: {
      style?: string;
      color?: string;
    };
    top?: {
      style?: string;
      color?: string;
    };
    bottom?: {
      style?: string;
      color?: string;
    };
    diagonal?: {
      style?: string;
      color?: string;
    };
    diagonalDown?: boolean;
    diagonalUp?: boolean;
    outline?: boolean;
  };
  fill?: {
    type?: string; // Currently only 'pattern' is implemented. Non-implemented option is 'gradient'
    patternType?: string; //§18.18.55 ST_PatternType (Pattern Type) see https://github.com/natergj/excel4node/blob/master/source/lib/types/fillPattern.js
    bgColor?: string; // HTML style hex value. defaults to black
    fgColor?: string; // HTML style hex value. defaults to black.
  };
  numberFormat?: string; // §18.8.30 numFmt (Number Format)
}

export interface WorkBookOpts {
  debug?: boolean;
  jszip?: { compression: "DEFLATE" | string };
  defaultFont?: {
    bold?: boolean;
    charset?: number;
    color?: string;
    condense?: boolean;
    extend?: boolean;
    family?: string;
    italics?: boolean;
    name?: string;
    outline?: boolean;
    scheme?: string; // §18.18.33 ST_FontScheme (Font scheme Styles)
    shadow?: boolean;
    strike?: boolean;
    size?: number;
    underline?: boolean;
    vertAlign?: string; // §22.9.2.17 ST_VerticalAlignRun (Vertical Positioning Location)
  };
  dateFormat?: string;
  workbookView?: {
    activeTab?: number; // Specifies an unsignedInt that contains the index to the active sheet in this book view.
    autoFilterDateGrouping?: boolean; // Specifies a boolean value that indicates whether to group dates when presenting the user with filtering options in the user interface.
    firstSheet?: number; // Specifies the index to the first sheet in this book view.
    minimized?: boolean; // Specifies a boolean value that indicates whether the workbook window is minimized.
    showHorizontalScroll?: boolean; // Specifies a boolean value that indicates whether to display the horizontal scroll bar in the user interface.
    showSheetTabs?: boolean; // Specifies a boolean value that indicates whether to display the sheet tabs in the user interface.
    showVerticalScroll?: boolean; // Specifies a boolean value that indicates whether to display the vertical scroll bar.
    tabRatio?: number; // Specifies ratio between the workbook tabs bar and the horizontal scroll bar.
    visibility?: "hidden" | "veryHidden" | "visible"; // Specifies visible state of the workbook window. ('hidden', 'veryHidden', 'visible') (§number8.number8.89)
    windowHeight?: number; // Specifies the height of the workbook window. The unit of measurement for this value is twips.
    windowWidth?: number; // Specifies the width of the workbook window. The unit of measurement for this value is twips..
    xWindow?: number; // Specifies the X coordinate for the upper left corner of the workbook window. The unit of measurement for this value is twips.
    yWindow?: number; //
  };
}

export interface WorkSheetStyle {
  margins?: {
    // Accepts a number in Inches
    bottom?: number;
    footer?: number;
    header?: number;
    left?: number;
    right?: number;
    top?: number;
  };
  printOptions?: {
    centerHorizontal?: boolean;
    centerVertical?: boolean;
    printGridLines?: boolean;
    printHeadings?: boolean;
  };
  headerFooter?: {
    // Set Header and Footer strings and options. See  https://poi.apache.org/apidocs/org/apache/poi/xssf/usermodel/extensions/XSSFHeaderFooter.html     i.e. '&L&A&C&BCompany, Inc. Confidential&B&RPage &P of &N'
    evenFooter?: string;
    evenHeader?: string;
    firstFooter?: string;
    firstHeader?: string;
    oddFooter?: string;
    oddHeader?: string;
    alignWithMargins?: boolean;
    differentFirst?: boolean;
    differentOddEven?: boolean;
    scaleWithDoc?: boolean;
  };
  pageSetup?: {
    blackAndWhite?: boolean;
    cellComments?: "none" | "asDisplayed" | "atEnd";
    copies?: number;
    draft?: boolean;
    errors?: "displayed" | "blank" | "dash" | "NA";
    firstPageNumber?: number;
    fitToHeight?: number; // Number of vertical pages to fit to
    fitToWidth?: number; // Number of horizontal pages to fit to
    horizontalDpi?: number;
    orientation?: "default" | "portrait" | "landscape";
    pageOrder?: "downThenOver" | "overThenDown";
    paperHeight?: string; // Value must a positive Float immediately followed by unit of measure from list mm, cm, in, pt, pc, pi. i.e. '10.5cm'
    paperSize?: number; // see https://github.com/natergj/excel4node/blob/master/source/lib/types/paperSize.js for all types and descriptions of types. setting paperSize overrides paperHeight and paperWidth settings
    paperWidth?: string; // Value must a positive Float immediately followed by unit of measure from list mm, cm, in, pt, pc, pi. i.e. '10.5cm'
    scale?: number;
    useFirstPageNumber?: boolean;
    usePrinterDefaults?: boolean;
    verticalDpi?: number;
  };
  sheetView?: {
    pane?: {
      // Note. Calling .freeze() on a row or column will adjust these values
      activePane?: "bottomLeft" | "bottomRight" | "topLeft" | "topRight";
      state?: "split" | "frozen" | "frozenSplit"; // one of 'split', 'frozen', 'frozenSplit'
      topLeftCell?: string; // i.e. 'A1'
      xSplit?: number; // Horizontal position of the split, in 1/20th of a point; 0 (zero) if none. If the pane is frozen, this value indicates the number of columns visible in the top pane.
      ySplit?: number; // Vertical position of the split, in 1/20th of a point; 0 (zero) if none. If the pane is frozen, this value indicates the number of rows visible in the left pane.
    };
    rightToLeft?: boolean; // Flag indicating whether the sheet is in 'right to left' display mode. When in this mode, Column A is on the far right, Column B ;is one column left of Column A, and so on. Also, information in cells is displayed in the Right to Left format.
    showGridLines?: boolean; // Flag indicating whether the sheet should have gridlines enabled or disabled during view
    zoomScale?: number; // Defaults to 100
    zoomScaleNormal?: number; // Defaults to 100
    zoomScalePageLayoutView?: number; // Defaults to 100
  };
  sheetFormat?: {
    baseColWidth?: number; // Defaults to 10. Specifies the number of characters of the maximum digit width of the normal style's font. This value does not include margin padding or extra padding for gridlines. It is only the number of characters.,
    defaultColWidth?: number;
    defaultRowHeight?: number;
    thickBottom?: boolean; // 'True' if rows have a thick bottom border by default.
    thickTop?: boolean; // 'True' if rows have a thick top border by default.
  };
  sheetProtection?: {
    // same as "Protect Sheet" in Review tab of Excel
    autoFilter?: boolean; // True means that that user will be unable to modify this setting
    deleteColumns?: boolean;
    deleteRows?: boolean;
    formatCells?: boolean;
    formatColumns?: boolean;
    formatRows?: boolean;
    insertColumns?: boolean;
    insertHyperlinks?: boolean;
    insertRows?: boolean;
    objects?: boolean;
    password?: string;
    pivotTables?: boolean;
    scenarios?: boolean;
    selectLockedCells?: boolean;
    selectUnlockedCells?: boolean;
    sheet?: boolean;
    sort?: boolean;
  };
  outline?: {
    summaryBelow?: boolean; // Flag indicating whether summary rows appear below detail in an outline, when applying an outline/grouping.
    summaryRight?: boolean; // Flag indicating whether summary columns appear to the right of detail in an outline, when applying an outline/grouping.
  };
  disableRowSpansOptimization?: boolean; // Flag indicating whether to remove the "spans" attribute on row definitions. Including spans in an optimization for Excel file readers but is not necessary,
  hidden?: boolean; // Flag indicating whether to not hide the worksheet within the workbook.
}

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
