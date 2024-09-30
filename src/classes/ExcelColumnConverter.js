"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var ExcelColumnConverter = /** @class */ (function () {
    function ExcelColumnConverter() {
    }
    ExcelColumnConverter.cellRefToIndex = function (reference) {
        var _a, _b;
        var match = (_a = reference.match(/^([A-Z]+)(\d+)$/)) !== null && _a !== void 0 ? _a : [];
        if (match.length < 3)
            throw new Error("Invalid cell reference '".concat(reference, "'"));
        var _c = match !== null && match !== void 0 ? match : ['', ''], columnLetter = _c[1], rowNumber = _c[2];
        var rowIndex = Number(rowNumber);
        var columnIndex = 0;
        if (this.columnToNumberMap.has(columnLetter)) {
            columnIndex = (_b = this.columnToNumberMap.get(columnLetter)) !== null && _b !== void 0 ? _b : -1;
        }
        else {
            columnIndex = 0;
            for (var i = 0; i < columnLetter.length; i++) {
                columnIndex = columnIndex * 26 + (columnLetter.charCodeAt(i) - 'A'.charCodeAt(0) + 1);
            }
            this.columnToNumberMap.set(columnLetter, columnIndex);
            if (!this.numberToColumnMap.has(columnIndex)) {
                this.numberToColumnMap.set(columnIndex, columnLetter);
            }
        }
        return { RowIndex: rowIndex, ColumnIndex: columnIndex };
    };
    ExcelColumnConverter.numberToColumn = function (number) {
        if (this.numberToColumnMap.has(number)) {
            return this.numberToColumnMap.get(number);
        }
        var column = '';
        while (number > 0) {
            number--; // Adjust for 0-indexing
            var remainder = number % 26;
            column = String.fromCharCode(remainder + 'A'.charCodeAt(0)) + column;
            number = Math.floor(number / 26);
        }
        this.numberToColumnMap.set(number, column);
        if (!this.columnToNumberMap.has(column)) {
            this.columnToNumberMap.set(column, number);
        }
        return column;
    };
    ExcelColumnConverter.columnToNumberMap = new Map();
    ExcelColumnConverter.numberToColumnMap = new Map();
    return ExcelColumnConverter;
}());
exports.default = ExcelColumnConverter;
