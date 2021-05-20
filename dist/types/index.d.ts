export interface CellStyle {
    alignment?: {
        horizontal?: "center" | "centerContinuous" | "distributed" | "fill" | "general" | "justify" | "left" | "right";
        indent?: number;
        justifyLastLine?: boolean;
        readingOrder?: "contextDependent" | "leftToRight" | "rightToLeft";
        relativeIndent?: number;
        shrinkToFit?: boolean;
        textRotation?: number;
        vertical?: "bottom" | "center" | "distributed" | "justify" | "top";
        wrapText?: boolean;
    };
    font?: {
        bold?: boolean;
        charset?: number;
        color?: string;
        condense?: boolean;
        extend?: boolean;
        family?: string;
        italics?: boolean;
        name?: string;
        outline?: boolean;
        scheme?: string;
        shadow?: boolean;
        strike?: boolean;
        size?: number;
        underline?: boolean;
        vertAlign?: string;
    };
    border?: {
        left?: {
            style?: string;
            color?: string;
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
        type?: string;
        patternType?: string;
        bgColor?: string;
        fgColor?: string;
    };
    numberFormat?: string;
}
export interface WorkBookOpts {
    debug?: boolean;
    jszip?: {
        compression: "DEFLATE" | string;
    };
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
        scheme?: string;
        shadow?: boolean;
        strike?: boolean;
        size?: number;
        underline?: boolean;
        vertAlign?: string;
    };
    dateFormat?: string;
    workbookView?: {
        activeTab?: number;
        autoFilterDateGrouping?: boolean;
        firstSheet?: number;
        minimized?: boolean;
        showHorizontalScroll?: boolean;
        showSheetTabs?: boolean;
        showVerticalScroll?: boolean;
        tabRatio?: number;
        visibility?: "hidden" | "veryHidden" | "visible";
        windowHeight?: number;
        windowWidth?: number;
        xWindow?: number;
        yWindow?: number;
    };
}
export interface WorkSheetStyle {
    margins?: {
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
        fitToHeight?: number;
        fitToWidth?: number;
        horizontalDpi?: number;
        orientation?: "default" | "portrait" | "landscape";
        pageOrder?: "downThenOver" | "overThenDown";
        paperHeight?: string;
        paperSize?: number;
        paperWidth?: string;
        scale?: number;
        useFirstPageNumber?: boolean;
        usePrinterDefaults?: boolean;
        verticalDpi?: number;
    };
    sheetView?: {
        pane?: {
            activePane?: "bottomLeft" | "bottomRight" | "topLeft" | "topRight";
            state?: "split" | "frozen" | "frozenSplit";
            topLeftCell?: string;
            xSplit?: number;
            ySplit?: number;
        };
        rightToLeft?: boolean;
        showGridLines?: boolean;
        zoomScale?: number;
        zoomScaleNormal?: number;
        zoomScalePageLayoutView?: number;
    };
    sheetFormat?: {
        baseColWidth?: number;
        defaultColWidth?: number;
        defaultRowHeight?: number;
        thickBottom?: boolean;
        thickTop?: boolean;
    };
    sheetProtection?: {
        autoFilter?: boolean;
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
        summaryBelow?: boolean;
        summaryRight?: boolean;
    };
    disableRowSpansOptimization?: boolean;
    hidden?: boolean;
}
export interface Cell {
    text: string;
    type?: CellType;
    rowSpan?: number;
    colSpan?: number;
    style?: CellStyle;
    childCells?: Cell[][];
}
export declare enum CellType {
    number = "number",
    string = "string",
    image = "image",
    date = "date",
    link = "link"
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
