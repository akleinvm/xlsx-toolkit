"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var ExcelWorkbook = /** @class */ (function () {
    function ExcelWorkbook() {
    }
    ExcelWorkbook.prototype.fromXML = function (xmlString) {
        this.xmlDocument = new DOMParser().parseFromString(xmlString, "text/xml");
        var sheets = this.xmlDocument.getElementsByTagName('sheet');
        this.sheets = [];
        for (var i = 0; i < sheets.length; i++) {
            var sheet = sheets[i];
            var sheetName = sheet.getAttribute('name');
            if (!sheetName)
                throw new Error('A sheetName is null or invalid');
            this.sheets[i] = sheetName;
        }
    };
    ExcelWorkbook.prototype.toString = function () {
        return new XMLSerializer().serializeToString(this.xmlDocument);
    };
    return ExcelWorkbook;
}());
exports.default = ExcelWorkbook;
