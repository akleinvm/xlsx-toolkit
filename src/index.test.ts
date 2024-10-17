import { expect, expectTypeOf, test } from "vitest";
import { JSDOM } from 'jsdom';
import ExcelDocument from "./index";
import {readFileSync} from 'fs';
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
  const file = readFileSync('./test/Project 1.xlsx');
  
  workbook = new ExcelDocument();
  await workbook.loadXLSX(file)
});

test("Read XLSX worksheet", async () => {
  if(!workbook) throw new Error('workbook is null');
  worksheet = await workbook.getWorksheet(1);
  
  rangeValues = worksheet.getRangeValues();
});

test("Update XLSX worksheet - Fill up values", async () => {
  const file = readFileSync('./test/Comparison of Valve Dimension Table S_DDMMYYYY.xlsx');
  
  updatedWorkbook = new ExcelDocument();
  await updatedWorkbook.loadXLSX(file);

  const updatedWorksheet = await updatedWorkbook.getWorksheet(6);
  
  const copiedRangeValues = rangeValues.splice(5, rangeValues.length);

  for (let rowNo = 0; rowNo < copiedRangeValues.length; rowNo++) {
    const row = copiedRangeValues[rowNo];

    for (let columnNo = 0; columnNo < row.length; columnNo++) {
      const column = row[columnNo];
      if(!column) continue;

      updatedWorksheet.updateCell({Value: column, Format: {Type: 's', Style: null}}, rowNo + 6, columnNo + 1);
    }
  }
});

test("Update XLSX worksheet - Fillup formulas", async () => {
  const updatedWorksheet = await updatedWorkbook.getWorksheet(4);
  const rangeFormulas = updatedWorksheet.getRangeFormulas();
  const copiedRangeFormulas = rangeFormulas.splice(5, 6)[0];
  console.log(copiedRangeFormulas);

  for (let rowNo = 0; rowNo < 10; rowNo++) {
    const row = copiedRangeFormulas;

    for (let columnNo = 0; columnNo < row.length; columnNo++) {
      const column = row[columnNo];
      if(!column) continue;
      const updatedColumn = column.replace('6', rowNo.toString());
      console.log(updatedColumn);

      updatedWorksheet.updateCell({Formula: updatedColumn, Format: {Type: 's', Style: null}}, rowNo + 5, columnNo + 1);
    }
  }
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

test("Generate sharedStrings file", async () => {
  
    const arrayBuffer = workbook.getSharedString();
  
    fs.writeFile('./test/output/sharedStrings.txt', arrayBuffer, (error) => {
      if (error) {
        console.error('Error saving the file:', error);
      } else {
          console.log('File saved successfully!');
      }
    })

});

test("Generate stringValues file", async () => {
  
    const arrayBuffer = JSON.stringify(worksheet.getRangeValues());
  
    fs.writeFile('./test/output/rangeValues.txt', arrayBuffer, (error) => {
      if (error) {
        console.error('Error saving the file:', error);
      } else {
          console.log('File saved successfully!');
      }
    })

})