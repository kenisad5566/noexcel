[![NPM version 1.0.1](https://www.npmjs.com/package/noexcel)
[![License](https://img.shields.io/badge/License-MIT-brightgreen.svg)](https://opensource.org/licenses/MIT)

# node-excel

A full featured xlsx file generation library allowing for the creation of advanced Excel files without consideration of cell detail.
With this module, it can export excel easily only focus on your work and data. [Some example](https://github.com/kenisad5566/node-excel/tree/master/example).

### Quick start

```javascript
npm i noexcel
```

### Basic Usage

A simple excel example like this

![image](https://github.com/kenisad5566/node-excel/blob/master/example/png/simple.png)

```javascript
import Excel  from "node-excel";
import {Cell}  from "node-excel/types";
const path = require("path");

// Create a new Instance of NodeExcel class
const excel = new Excel();

// Add a worksheet
const sheetName1 = "sheet1";
excel.addWorkSheet(sheetName1);

// Set save file path if need
const exportPath = path.join(__dirname, "../tmp");
excel.setSavePath(exportPath);

// set file name
const fileName = 'test'
excel.setFileName(fileName)

// render simple data

const simpleData = [
    [{ text: "s/n" }, { text: "name" }, { text: "age" }, { text: "position" }],
    [{ text: "1" }, { text: "ming" }, { text: "15" }, { text: "monitor" }],
    [{ text: "2" }, { text: "hua" }, { text: "14" }, { text: "commissary" }],
    [{ text: "3" }, { text: "ai" }, { text: "14" }, { text: "supervisor" }],
  ] as Cell[][];

// render simpleData
await excel.render(simpleData);


// save file to exportPath
await excel.saveFile()

```

### ColSpan

A colspan excel example like this

![image](https://github.com/kenisad5566/node-excel/blob/master/example/png/colSpan.png)

```javascript
 const exportPath = path.join(__dirname, "../tmp");
  const fileName = "test";
  const sheetName = "student";

  const data = [
    [
      { text: "class", colSpan: 2 },
      { text: "name" },
      { text: "age" },
      { text: "position" },
    ],
    [
      {
        text: "class 1",
        colSpan: 2,
      },
      { text: "hua" },
      { text: "14" },
      { text: "commissary" },
    ],
    [
      { text: "class 2", colSpan: 2 },
      { text: "hua" },
      { text: "14" },
      { text: "commissary" },
    ],
    [
      {
        text: "class 3",
        colSpan: 2,
        rowSpan: 2,
        childCells: [
          [{ text: "ai" }, { text: "13" }, { text: "supervisor" }],
          [{ text: "ai" }, { text: "13" }, { text: "supervisor" }],
        ],
      },
    ],
  ] as Cell[][];

  const excel = new Excel();
  excel.addWorkSheet(sheetName).setSavePath(exportPath).setFileName(fileName);
  await excel.render(data);
  await excel.saveFile();

```

### RowSpan

A rowspan excel example like this

![image](https://github.com/kenisad5566/node-excel/blob/master/example/png/rowSpan.png)

```javascript
  const exportPath = path.join(__dirname, "../tmp");
  const fileName = "test";
  const sheetName = "student";

  const data = [
    [
      { text: "class" },
      { text: "name" },
      { text: "age" },
      { text: "position" },
    ],
    [
      {
        text: "class one",
        rowSpan: 3,
        childCells: [
          [{ text: "ming" }, { text: "15" }, { text: "monitor" }],
          [{ text: "ai" }, { text: "15" }, { text: "commissary" }],
          [{ text: "ai" }, { text: "15" }, { text: "supervisor" }],
        ],
      },
    ],
    [{ text: "2" }, { text: "hua" }, { text: "14" }, { text: "commissary" }],
    [{ text: "3" }, { text: "ai" }, { text: "13" }, { text: "supervisor" }],
  ] as Cell[][];

  const excel = new Excel();
  excel.addWorkSheet(sheetName).setSavePath(exportPath).setFileName(fileName);
  await excel.render(data);
  await excel.saveFile();
```

### Cell style

Some other example with cell style like this

![image](https://github.com/kenisad5566/node-excel/blob/master/example/png/styles.png)

```javascript
  const exportPath = path.join(__dirname, "../tmp");
  const fileName = "test";
  const sheetName = "student";

  const data = [
    [
      { text: "s/n" },
      { text: "name" },
      { text: "age" },
      { text: "position" },
      { text: "date" },
      { text: "link" },
    ],
    [
      { text: "1", style: { font: { bold: true } } },
      { text: "ming" },
      { text: 15, type: "number" },
      { text: "monitor" },
      { text: new Date(), type: "date", style: { numberFormat: "yyyy-mm-dd" } },
      { text: "http://www.google.com", type: "link" },
    ],
    [
      { text: "2" },
      { text: "hua", style: { font: { size: 14 } } },
      { text: 14, type: "number" },
      { text: "commissary" },
      { text: new Date(), type: "date", style: { numberFormat: "yyyy-mm-dd" } },
      { text: "http://www.google.com", type: "string" },
    ],
    [
      { text: "3" },
      { text: "ai" },
      { text: 13 },
      { text: "supervisor" },
      { text: new Date(), type: "date", style: { numberFormat: "yyyy-mm-dd" } },
      {
        text: "http://www.google.com",
        type: "link",
        style: { font: { underline: true, bold: true, color: "black" } },
      },
    ],
    [
      { text: "4" },
      { text: "ai" },
      { text: 14 },
      { text: "supervisor" },
      { text: new Date(), type: "date", style: { numberFormat: "yyyy-mm-dd" } },
      {
        text: "http://www.cnlogo8.com/d/file/2021-05-20/97517b732413c71921c3a55726f4f299.png",
        type: "image",
      },
    ],
  ] as Cell[][];

  const excel = new Excel();
  excel.addWorkSheet(sheetName).setSavePath(exportPath).setFileName(fileName);
  await excel.render(data);
  excel.setRowHeight(5, 300).setColWidth(6, 70);

  await excel.saveFile();
```

### Options and Cell types

WorkbookOptions ,WorkSheetOptions and cell style can see https://github.com/natergj/excel4node/blob/master/README.md

```javascript
 Cell {
  text: string; // cell value, if type is image, text should be url
  type?: CellType; // cellType, has five value, see below
  rowSpan?: number; // default 1, the cell row merge number
  colSpan?: number; // default 1, the cell column merge number
  style?: CellStyle; // cell style, see types below or Style Objects in  https://github.com/natergj/excel4node/blob/master/README.md
  childCells?: Cell[][]; // if rowSpan >1 , childCells maybe not empty
}

 CellType {
    number = "number",
    string = "string",
    image = "image",
    date = "date",
    link = "link"
}

CellStyle {
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

 WorkBookOpts {
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

 WorkSheetStyle {
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

```

### worksheet and filename

Add wordSheet with options and set sheet name

```javascript
excel.addWorkSheet(sheetName, sheetOptions);
```

Set save file path and file name

```javascript
excel.setFileName(fineName);
excel.setSuffix(suffix); // suffix should be .xlsx and .xls
```

### Rows and Columns

Set custom widths and heights of columns/rows

```javascript
excel.setRowHeight(5, 300).setColWidth(6, 70);
```

Set rows and/or columns to create a frozen pane with an optionall scrollTo

```javascript
excel.setRowFreeze(1, 1);
excel.setColFreeze(1, 1);
```

Hide a row or column

```javascript
excel.setRowHide(1, 1);
excel.setColHide(1, 1);
```

### Save as file or read as buffer

It can be save as file or write to buffer

```javascript
const filePath = await excel.saveFile(); // save file and return the file path
const buffer = await excel.writeToBuffer(); // you can return this buffer to http.response, some it can be export by browser
```

### Parse excel file

It can also parse a excel file to array data

```javascript
const data = await excel.readExcel(filePath, sheetIndex); // save file and return the file path
```
