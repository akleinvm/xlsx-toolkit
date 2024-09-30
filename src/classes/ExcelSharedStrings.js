"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
var ExcelSharedStrings = /** @class */ (function () {
    function ExcelSharedStrings() {
    }
    ExcelSharedStrings.prototype.fromXML = function (xmlString) {
        var _a;
        this.xmlDocument = new DOMParser().parseFromString(xmlString, 'text/xml');
        this.sstElement = this.xmlDocument.getElementsByTagName('sst')[0];
        this.namespace = (_a = this.sstElement.getAttribute('xmlns')) !== null && _a !== void 0 ? _a : '';
        this.stringsMap = new Map();
        this.numbersMap = new Map();
        var siElements = this.sstElement.getElementsByTagName('si');
        for (var i = 0; i < siElements.length; i++) {
            var tElement = siElements[i].getElementsByTagName('t')[0];
            if (tElement && tElement.textContent) {
                this.stringsMap.set(tElement.textContent, i);
                this.numbersMap.set(i, tElement.textContent);
            }
        }
    };
    ExcelSharedStrings.prototype.toString = function () {
        this.sstElement.setAttribute('uniqueCount', this.stringsMap.size.toString());
        return new XMLSerializer().serializeToString(this.xmlDocument);
    };
    ExcelSharedStrings.prototype.getStringIndex = function (string) {
        var stringIndex = this.stringsMap.get(string);
        if (!stringIndex) {
            stringIndex = this.stringsMap.size;
            this.stringsMap.set(string, this.stringsMap.size);
            this.numbersMap.set(this.stringsMap.size, string);
            var siElement = this.xmlDocument.createElementNS(this.namespace, 'si');
            var tElement = this.xmlDocument.createElementNS(this.namespace, 't');
            tElement.textContent = string;
            siElement.appendChild(tElement);
            this.sstElement.appendChild(siElement);
        }
        return stringIndex;
    };
    ExcelSharedStrings.prototype.getIndexString = function (index) {
        var _a;
        return (_a = this.numbersMap.get(index)) !== null && _a !== void 0 ? _a : '';
    };
    return ExcelSharedStrings;
}());
exports.default = ExcelSharedStrings;
