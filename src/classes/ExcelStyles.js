"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var ExcelStyles = /** @class */ (function () {
    function ExcelStyles() {
    }
    ExcelStyles.prototype.fromXML = function (xmlString) {
        var _a;
        this.xmlDocument = new DOMParser().parseFromString(xmlString, "text/xml");
        this.styleSheetElement = this.xmlDocument.getElementsByTagName('styleSheet')[0];
        this.namespace = (_a = this.styleSheetElement.getAttribute('xmlns')) !== null && _a !== void 0 ? _a : "";
        this.cellFormatsElement = this.styleSheetElement.getElementsByTagName('cellXfs')[0];
        var formats = this.cellFormatsElement.getElementsByTagName('xf');
        this.cellFormatArray = [];
        for (var i = 0; i < formats.length; i++) {
            var format = formats[i];
            var formatId = Number(format.getAttribute('numFmtId'));
            this.cellFormatArray.push(formatId);
        }
    };
    ExcelStyles.prototype.getFormatIndex = function (formatId) {
        var formatIndex = this.cellFormatArray.indexOf(formatId).toString();
        if (formatIndex === '-1') {
            formatIndex = this.cellFormatArray.length.toString();
            this.cellFormatArray.push(formatId);
            var xfElement = this.xmlDocument.createElementNS(this.namespace, 'xf');
            xfElement.setAttribute('numFmtId', formatId.toString());
            this.cellFormatsElement.appendChild(xfElement);
        }
        return formatIndex;
    };
    ExcelStyles.prototype.toString = function () {
        return new XMLSerializer().serializeToString(this.xmlDocument);
    };
    return ExcelStyles;
}());
exports.default = ExcelStyles;
