"use strict";
var __awaiter = (this && this.__awaiter) || function (thisArg, _arguments, P, generator) {
    function adopt(value) { return value instanceof P ? value : new P(function (resolve) { resolve(value); }); }
    return new (P || (P = Promise))(function (resolve, reject) {
        function fulfilled(value) { try { step(generator.next(value)); } catch (e) { reject(e); } }
        function rejected(value) { try { step(generator["throw"](value)); } catch (e) { reject(e); } }
        function step(result) { result.done ? resolve(result.value) : adopt(result.value).then(fulfilled, rejected); }
        step((generator = generator.apply(thisArg, _arguments || [])).next());
    });
};
var __generator = (this && this.__generator) || function (thisArg, body) {
    var _ = { label: 0, sent: function() { if (t[0] & 1) throw t[1]; return t[1]; }, trys: [], ops: [] }, f, y, t, g = Object.create((typeof Iterator === "function" ? Iterator : Object).prototype);
    return g.next = verb(0), g["throw"] = verb(1), g["return"] = verb(2), typeof Symbol === "function" && (g[Symbol.iterator] = function() { return this; }), g;
    function verb(n) { return function (v) { return step([n, v]); }; }
    function step(op) {
        if (f) throw new TypeError("Generator is already executing.");
        while (g && (g = 0, op[0] && (_ = 0)), _) try {
            if (f = 1, y && (t = op[0] & 2 ? y["return"] : op[0] ? y["throw"] || ((t = y["return"]) && t.call(y), 0) : y.next) && !(t = t.call(y, op[1])).done) return t;
            if (y = 0, t) op = [op[0] & 2, t.value];
            switch (op[0]) {
                case 0: case 1: t = op; break;
                case 4: _.label++; return { value: op[1], done: false };
                case 5: _.label++; y = op[1]; op = [0]; continue;
                case 7: op = _.ops.pop(); _.trys.pop(); continue;
                default:
                    if (!(t = _.trys, t = t.length > 0 && t[t.length - 1]) && (op[0] === 6 || op[0] === 2)) { _ = 0; continue; }
                    if (op[0] === 3 && (!t || (op[1] > t[0] && op[1] < t[3]))) { _.label = op[1]; break; }
                    if (op[0] === 6 && _.label < t[1]) { _.label = t[1]; t = op; break; }
                    if (t && _.label < t[2]) { _.label = t[2]; _.ops.push(op); break; }
                    if (t[2]) _.ops.pop();
                    _.trys.pop(); continue;
            }
            op = body.call(thisArg, _);
        } catch (e) { op = [6, e]; y = 0; } finally { f = t = 0; }
        if (op[0] & 5) throw op[1]; return { value: op[0] ? op[1] : void 0, done: true };
    }
};
Object.defineProperty(exports, "__esModule", { value: true });
var jszip_1 = require("jszip");
var ExcelStyles_1 = require("./ExcelStyles");
var ExcelSharedStrings_1 = require("./ExcelSharedStrings");
var ExcelWorkbook_1 = require("./ExcelWorkbook");
var ExcelWorksheet_1 = require("./ExcelWorksheet");
var ExcelDocument = /** @class */ (function () {
    function ExcelDocument() {
    }
    ExcelDocument.prototype.loadXLSX = function (arrayBuffer) {
        return __awaiter(this, void 0, void 0, function () {
            var zip, fileWorkbook, xmlWorkbook, fileSharedStrings, xmlSharedStrings, fileStyles, xmlStyles;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        console.log('Extracting xlsx document...');
                        zip = new jszip_1.default();
                        return [4 /*yield*/, zip.loadAsync(arrayBuffer)];
                    case 1:
                        _a.sent();
                        console.log('Retrieving files...');
                        this.zipFiles = zip.folder("");
                        if (!this.zipFiles)
                            throw new Error('no file found');
                        this.zipFiles.remove('xl/calcChain.xml');
                        console.log('Parsing workbook.xml');
                        this.workbook = new ExcelWorkbook_1.default();
                        fileWorkbook = this.zipFiles.file('xl/workbook.xml');
                        if (!fileWorkbook)
                            throw new Error('workbook not found');
                        return [4 /*yield*/, fileWorkbook.async('binarystring')];
                    case 2:
                        xmlWorkbook = _a.sent();
                        if (!xmlWorkbook)
                            throw new Error('workbook is null or undefined');
                        this.workbook.fromXML(xmlWorkbook);
                        console.log('Parsing sharedStrings.xml...');
                        this.sharedStrings = new ExcelSharedStrings_1.default();
                        fileSharedStrings = this.zipFiles.file('xl/sharedStrings.xml');
                        if (!fileSharedStrings)
                            throw new Error('sharedStrings not found');
                        return [4 /*yield*/, fileSharedStrings.async('binarystring')];
                    case 3:
                        xmlSharedStrings = _a.sent();
                        if (!xmlSharedStrings)
                            throw new Error('sharedStrings is null or undefined');
                        this.sharedStrings.fromXML(xmlSharedStrings);
                        console.log('Parsing styles.xml...');
                        this.styles = new ExcelStyles_1.default();
                        fileStyles = this.zipFiles.file("xl/styles.xml");
                        if (!fileStyles)
                            throw new Error('styles not found');
                        return [4 /*yield*/, fileStyles.async('binarystring')];
                    case 4:
                        xmlStyles = _a.sent();
                        if (!xmlStyles)
                            throw new Error('styles is null or undefined');
                        this.styles.fromXML(xmlStyles);
                        this.worksheets = new Map();
                        return [2 /*return*/];
                }
            });
        });
    };
    ExcelDocument.prototype.getWorksheet = function (sheetNo) {
        return __awaiter(this, void 0, void 0, function () {
            var fileWorksheet, xmlWorksheet, worksheet;
            return __generator(this, function (_a) {
                switch (_a.label) {
                    case 0:
                        if (!this.zipFiles)
                            throw new Error('no file found');
                        console.log("Parsing sheet".concat(sheetNo, ".xml..."));
                        fileWorksheet = this.zipFiles.file("xl/worksheets/sheet".concat(sheetNo, ".xml"));
                        if (!fileWorksheet)
                            throw new Error('worksheet not found');
                        return [4 /*yield*/, fileWorksheet.async('binarystring')];
                    case 1:
                        xmlWorksheet = _a.sent();
                        if (!xmlWorksheet)
                            throw new Error('worksheet is null or undefined');
                        worksheet = new ExcelWorksheet_1.default(this.sharedStrings);
                        worksheet.fromXML(xmlWorksheet);
                        this.worksheets.set(sheetNo, worksheet);
                        /*
                        for (let i = 0; i < this.workbook.sheets.length; i++) {
                            console.log(`Parsing sheet${i + 1}.xml...`);
                            const xmlWorksheet = this.files.get(`xl/worksheets/sheet${i + 1}.xml`);
                            if(!xmlWorksheet) throw new Error('worksheet is null or undefined');
                            const worksheet = new ExcelWorksheet(this.sharedStrings);
                            //const xmlWorksheet = bufferToString(bufferWorksheet);
                            worksheet.fromXML(xmlWorksheet);
                
                            const worksheetName = this.workbook.sheets[i].getAttribute('name');
                            if(!worksheetName) throw new Error('worksheetName is null or undefined');
                            this.worksheets.set(worksheetName, worksheet);
                        }
                
                
                        const worksheet = this.worksheets.get(sheetName);
                        if(!worksheet) throw new Error(`${sheetName} can't be found`);
                */
                        return [2 /*return*/, worksheet];
                }
            });
        });
    };
    ExcelDocument.prototype.saveXLSX = function () {
        return __awaiter(this, void 0, void 0, function () {
            var _i, _a, _b, sheetNo, worksheet, arrayBuffer;
            var _c;
            return __generator(this, function (_d) {
                switch (_d.label) {
                    case 0:
                        if (!this.zipFiles)
                            throw new Error('no file found');
                        console.log('Updating sharedString.xml...');
                        this.zipFiles.file("xl/sharedStrings.xml", this.sharedStrings.toString());
                        console.log('Updating styles.xml...');
                        this.zipFiles.file("xl/styles.xml", this.styles.toString());
                        console.log('Updating worksheets.xml...');
                        for (_i = 0, _a = this.worksheets; _i < _a.length; _i++) {
                            _b = _a[_i], sheetNo = _b[0], worksheet = _b[1];
                            (_c = this.zipFiles) === null || _c === void 0 ? void 0 : _c.file("xl/worksheets/sheet".concat(sheetNo, ".xml"), worksheet.toString());
                        }
                        return [4 /*yield*/, this.zipFiles.generateAsync({ type: 'arraybuffer', compression: "DEFLATE", compressionOptions: { level: 9 } })];
                    case 1:
                        arrayBuffer = _d.sent();
                        return [2 /*return*/, arrayBuffer];
                }
            });
        });
    };
    return ExcelDocument;
}());
exports.default = ExcelDocument;
function bufferToString(buffer) {
    var bytes = new Uint8Array(buffer);
    var binaryString = '';
    for (var i = 0; i < bytes.length; i++) {
        binaryString += String.fromCharCode(bytes[i]);
    }
    return binaryString;
}
