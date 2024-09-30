"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var ExcelColumnConverter_1 = require("./ExcelColumnConverter");
var ExcelWorksheet = /** @class */ (function () {
    function ExcelWorksheet(sharedStrings) {
        this.sharedStrings = sharedStrings;
    }
    ExcelWorksheet.prototype.fromXML = function (xmlString) {
        var _a, _b;
        this.xmlDocument = new DOMParser().parseFromString(xmlString, "text/xml");
        this.worksheetElement = this.xmlDocument.getElementsByTagName('worksheet')[0];
        this.namespace = (_a = this.worksheetElement.getAttribute('xmlns')) !== null && _a !== void 0 ? _a : "";
        this.sheetDataElement = this.xmlDocument.getElementsByTagName("sheetData")[0];
        this.rowsMap = new Map();
        var rows = this.sheetDataElement.getElementsByTagName('row');
        for (var i = 0; i < rows.length; i++) {
            var row = rows[i];
            this.rowsMap.set(Number(row.getAttribute('r')), row);
        }
        this.cellsMap = new Map();
        var cells = this.sheetDataElement.getElementsByTagName('c');
        for (var i = 0; i < cells.length; i++) {
            var cell = cells[i];
            this.cellsMap.set((_b = cell.getAttribute('r')) !== null && _b !== void 0 ? _b : '', cell);
        }
    };
    ExcelWorksheet.prototype.addCell = function (cell, rowNo, columnNo) {
        var _a, _b;
        var rowElement = this.rowsMap.get(rowNo);
        if (!rowElement) {
            rowElement = this.xmlDocument.createElementNS(this.namespace, 'row');
            rowElement.setAttribute('r', rowNo.toString());
            rowElement.setAttribute('spans', "".concat(1, ":").concat(columnNo));
            this.sheetDataElement.appendChild(rowElement);
            this.rowsMap.set(rowNo, rowElement);
        }
        else {
            var _c = (_b = (_a = rowElement.getAttribute('spans')) === null || _a === void 0 ? void 0 : _a.split(':')) !== null && _b !== void 0 ? _b : [columnNo, columnNo], minSpan = _c[0], maxSpan = _c[1];
            var spans = "".concat(Math.min(columnNo, Number(minSpan)).toString(), ":").concat(Math.max(columnNo, Number(maxSpan)));
            rowElement.setAttribute('spans', spans);
        }
        var cellReference = ExcelColumnConverter_1.default.numberToColumn(columnNo) + Number(rowNo);
        var cellElement = this.cellsMap.get(cellReference);
        if (!cellElement) {
            cellElement = this.xmlDocument.createElementNS(this.namespace, 'c');
            cellElement.setAttribute('r', cellReference);
        }
        var cellStyle = cell.Format.Style;
        if (!cellStyle)
            cellElement.removeAttribute('s');
        else
            cellElement.setAttribute('s', cellStyle);
        var cellType = cell.Format.Type;
        if (!cellType)
            cellElement.removeAttribute('t');
        else
            cellElement.setAttribute('t', cellType);
        var valueElement = this.xmlDocument.createElementNS(this.namespace, 'v');
        valueElement.textContent = cell.Value.toString();
        cellElement.replaceChildren(valueElement);
        console.log(cellElement.textContent);
        rowElement.appendChild(cellElement);
        this.cellsMap.set(cellReference, cellElement);
    };
    ExcelWorksheet.prototype.getRangeValues = function () {
        var output = [];
        for (var _i = 0, _a = this.cellsMap; _i < _a.length; _i++) {
            var _b = _a[_i], key = _b[0], cell = _b[1];
            var _c = ExcelColumnConverter_1.default.cellRefToIndex(key), RowIndex = _c.RowIndex, ColumnIndex = _c.ColumnIndex;
            var rowNo = RowIndex - 1;
            var columnNo = ColumnIndex - 1;
            var valueElement = cell.querySelector('v');
            if (!(valueElement === null || valueElement === void 0 ? void 0 : valueElement.textContent))
                continue;
            var cellValue = valueElement.textContent;
            var cellType = cell.getAttribute('t');
            if (cellType === 's')
                cellValue = this.sharedStrings.getIndexString(Number(cellValue));
            if (!output[rowNo])
                output[rowNo] = [];
            output[rowNo][columnNo] = cellValue;
        }
        return output;
    };
    ExcelWorksheet.prototype.toString = function () {
        return new XMLSerializer().serializeToString(this.xmlDocument);
    };
    return ExcelWorksheet;
}());
exports.default = ExcelWorksheet;
/*

  private getValues(rangeStart: string, rangeEnd: string): Array<Array<string>> {
    const startIndex = ExcelColumnConverter.cellRefToIndex(rangeStart);
    const minRowNo = startIndex.RowIndex;
    const minColumnNo = startIndex.ColumnIndex;

    const endIndex = ExcelColumnConverter.cellRefToIndex(rangeEnd);
    const maxRowNo = endIndex.RowIndex;
    const maxColumnNo = endIndex.ColumnIndex;

    const output = new Array<Array<string>>();
    for (let rowNo = minRowNo; rowNo <= maxRowNo; rowNo++) {
      const arrayRowNo = rowNo - minRowNo;
      output[arrayRowNo] = [];

      for (let columnNo = minColumnNo; columnNo <= maxColumnNo; columnNo++) {
        const currentColumnRef = ExcelColumnConverter.numberToColumn(columnNo);
        const currentCell = this.cellsMap.get(currentColumnRef + rowNo.toString());
        
        let cellValue = '';
        const cell = currentCell?.getElementsByTagName('v')[0];
        if(cell) {
          cellValue = cell.textContent ?? '';
          const cellType = currentCell?.getAttribute('t');
          if(cellType === 's') cellValue = this.sharedStrings.getIndexString(Number(cellValue));
        }
        
        const arrayColumnNo = columnNo - minColumnNo;
        output[arrayRowNo][arrayColumnNo] = cellValue;
      }
    }
    
    return output;
  }
    */
/*

public addRows(rows: Array<Array<CellObject>>, rowStartIndex: number, columnStartIndex: number): void {
  for(let i = 0; i < rows.length; i++) {
    const rowIndex = rowStartIndex + i;
    const row = rows[i];

    let rowElement = this.rowsMap.get(rowIndex);
    if(!rowElement) {
      rowElement = this.xmlDocument.createElementNS(this.namespace, 'row');
      rowElement.setAttribute('r', rowIndex.toString());
      rowElement.setAttribute('spans', `${columnStartIndex}:${columnStartIndex + row.length}`);
      this.sheetDataElement.appendChild(rowElement);
      this.rowsMap.set(rowIndex, rowElement);
    } else {
      const [minSpan, maxSpan] = rowElement.getAttribute('spans')?.split(':') ?? [];
      const spans = `${Math.min(columnStartIndex, Number(minSpan)).toString()}:${Math.max(columnStartIndex + row.length - 1, Number(maxSpan))}`;
      rowElement.setAttribute('spans', spans);
    }
    
    for(let j = 0; j < row.length; j++) {
      const columnIndex = columnStartIndex + j;
      
      const cellReference = ExcelColumnConverter.numberToColumn(columnIndex) + rowIndex;
      let cellElement = this.cellsMap.get(cellReference);
      if(!cellElement) {
        cellElement = this.xmlDocument.createElementNS(this.namespace, 'c');
        cellElement.setAttribute('r', cellReference);
      }
      
      const cellStyle = row[j].Format.Style;
      if(cellStyle) cellElement.setAttribute('s', cellStyle ?? "");

      const cellType = row[j].Format.Type;
      if(cellType) cellElement.setAttribute('t', cellType ?? "");
      
      const valueElement = this.xmlDocument.createElementNS(this.namespace, 'v');
      valueElement.textContent = row[j].Value.toString();
      cellElement.replaceChildren(valueElement);
      rowElement.appendChild(cellElement);
      
      this.cellsMap.set(cellReference, cellElement);
    }
  }
}*/ 
