import { expect, expectTypeOf, test } from "vitest";
import { JSDOM } from 'jsdom';
import ExcelDocument from "./index";
import {readFileSync} from 'fs';
import ExcelColumnConverter from "./classes/ExcelColumnConverter";
import fs from 'fs';
import ExcelWorksheet from "./classes/ExcelWorksheet";

let workbook: ExcelDocument;
let worksheet: ExcelWorksheet;

let updatedWorkbook: ExcelDocument;
let rangeValues: Array<Array<string>>;

const {window} = new JSDOM('');
global.DOMParser = window.DOMParser;
global.Document  = window.Document;
global.XMLSerializer = window.XMLSerializer;

test("Load XLSX file", async () => {
  const file = readFileSync('./test/test1.xlsx');
  
  workbook = new ExcelDocument();
  await workbook.loadXLSX(file)
});

test("Read XLSX worksheet", async () => {
  if(!workbook) throw new Error('workbook is null');
  worksheet = await workbook.getWorksheet(1);
  
  rangeValues = worksheet.getRangeValues();
});

test("Update XLSX worksheet", async () => {
  const file = readFileSync('./test/Book1.xlsx');
  
  updatedWorkbook = new ExcelDocument();
  await updatedWorkbook.loadXLSX(file);

  const updatedWorksheet = await updatedWorkbook.getWorksheet(1);

  for (let rowNo = 0; rowNo < rangeValues.length; rowNo++) {
    const row = rangeValues[rowNo];

    for (let columnNo = 0; columnNo < row.length; columnNo++) {
      const column = row[columnNo];

      updatedWorksheet.addCell({Value: column, Format: {Type: 's', Style: null}}, rowNo + 1, columnNo + 1);
    }
  }
  

  //worksheet.addRows([[{Value: 3234, Format: {Type: null, Style: null}}]], 9, 9);
  /*
  for (let i = 1; i < 20; i++) {
    for (let j = 1; j < 20; j++) {
      worksheet.addCell({Value: 696969696969, Format: {Type: null, Style: null}}, i, j);
    }
  }*/
  


});

test("Save XLSX file", async () => {

  const arrayBuffer = await updatedWorkbook.saveXLSX();
  if(!arrayBuffer) return;
  const buffer = Buffer.from(arrayBuffer);

  fs.writeFile('./test/output/test.xlsx', buffer, (error) => {
    if (error) {
      console.error('Error saving the file:', error);
    } else {
        console.log('File saved successfully!');
    }
  })

});