import { expect, expectTypeOf, test } from "vitest";
import { JSDOM } from 'jsdom';
import ExcelDocument from "./index";
import {readFileSync} from 'fs';
import fs from 'fs';
import ExcelWorksheet from "./classes/ExcelWorksheet";
import { CellObject } from "./types";
import ExcelColumnConverter from "./classes/ExcelColumnConverter";

let previousSourceWorkbook: ExcelDocument;
let previousSourceWorksheet: ExcelWorksheet;
let previousSourceRange: (CellObject | undefined)[][];

let latestSourceWorkbook: ExcelDocument;
let latestSourceWorksheet: ExcelWorksheet;
let latestSourceRange: (CellObject | undefined)[][];

let templateWorkbook: ExcelDocument;
let templateWorksheet1: ExcelWorksheet;
let templateWorksheet2: ExcelWorksheet;
let templateWorksheet3: ExcelWorksheet;
let templateWorksheet4: ExcelWorksheet;
let templateWorksheet5: ExcelWorksheet;
let templateWorksheet6: ExcelWorksheet;

const {window} = new JSDOM('');
global.DOMParser = window.DOMParser;
global.Document  = window.Document;
global.XMLSerializer = window.XMLSerializer;

test("Load XLSX file", async () => {
  const previousSourcefile = readFileSync('./test/Project 1.xlsx');
  
  previousSourceWorkbook = new ExcelDocument();
  await previousSourceWorkbook.loadXLSX(previousSourcefile)
  
  const latestSourcefile = readFileSync('./test/Project 2.xlsx');
  latestSourceWorkbook = new ExcelDocument();
  await latestSourceWorkbook.loadXLSX(latestSourcefile)
});

test("Read XLSX worksheet", async () => {
  if(!previousSourceWorkbook) throw new Error('workbook is null');
  previousSourceWorksheet = await previousSourceWorkbook.getWorksheet(1);
  previousSourceRange = previousSourceWorksheet.getRange("B6", "BJ9999999999"); //console.log(previousSourceRange);

  if(!latestSourceWorkbook) throw new Error('workbook is null');
  latestSourceWorksheet = await latestSourceWorkbook.getWorksheet(1);
  latestSourceRange = latestSourceWorksheet.getRange("B6", "BJ9999999999");
});

test("Update XLSX worksheet - Fill up values", async () => {
  const file = readFileSync('./test/Comparison of Valve Dimension Table S_DDMMYYYY.xlsx');
  
  templateWorkbook = new ExcelDocument();
  await templateWorkbook.loadXLSX(file);

  templateWorksheet6 = await templateWorkbook.getWorksheet(6);
  templateWorksheet6.updateRange(previousSourceRange, 'B6');
  
  templateWorksheet5 = await templateWorkbook.getWorksheet(5);
  templateWorksheet5.updateRange(latestSourceRange, 'B6');
});

test("Update XLSX worksheet - Fillup formulas", async () => {
  const updateFormula = async (workSheetNo: number): Promise<void> => {
    const templateWorksheet4 = await templateWorkbook.getWorksheet(workSheetNo);
    const template4Formulas = templateWorksheet4.getRange('B6', 'BJ6')[0]; //console.log(template4Formulas);
    for (let rowNo = 0; rowNo < latestSourceRange.length; rowNo++) {
      const row = template4Formulas;

      for (let columnNo = 0; columnNo < row.length; columnNo++) {
        const column = row[columnNo];
        if(!column?.formula) continue;
        const updatedColumn = column.formula.replaceAll('6', (rowNo + 6).toString());

        const cellRef = ExcelColumnConverter.numberToColumn(columnNo + 2) + Number(rowNo + 6);
        templateWorksheet4.updateCell({formula: updatedColumn, format: column.format}, cellRef);
      }
    }
  };

  await updateFormula(4);
  await updateFormula(3);
  await updateFormula(2);
  await updateFormula(1);
  
});

test("Save XLSX file", async () => {

  const arrayBuffer = await templateWorkbook.saveXLSX();
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
/*
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
  
    const arrayBuffer = JSON.stringify(worksheet.getRange("A1", "ZZZZZ999999999"));
  
    fs.writeFile('./test/output/rangeValues.txt', arrayBuffer, (error) => {
      if (error) {
        console.error('Error saving the file:', error);
      } else {
          console.log('File saved successfully!');
      }
    })

})*/