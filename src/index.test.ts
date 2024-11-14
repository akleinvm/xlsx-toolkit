/* eslint-disable no-debugger */
import { expect, expectTypeOf, test } from "vitest";
import { JSDOM } from 'jsdom';
import {ExcelDocument} from "./index";
import {readFileSync} from 'fs';
import fs from 'fs';
import ExcelWorksheet from "./classes/ExcelWorksheet";
import { CellObject } from "./types";
import ExcelColumnConverter from "./classes/ExcelColumnConverter";

const {window} = new JSDOM('');
global.DOMParser = window.DOMParser;
global.Document  = window.Document;
global.XMLSerializer = window.XMLSerializer;

let previousSourceWorkbook: ExcelDocument;
let previousSourceWorksheet: ExcelWorksheet;
let previousSourceRange: (CellObject | null)[][];

let latestSourceWorkbook: ExcelDocument;
let latestSourceWorksheet: ExcelWorksheet;
let latestSourceRange: (CellObject | null)[][];

let templateWorkbook: ExcelDocument;
let templateWorksheet6: ExcelWorksheet;
let templateWorksheet7: ExcelWorksheet;

test("Load XLSX file", async () => {
  const previousSourcefile = readFileSync('./test/1. Source File_OLD_CV_Basrah_2.0.xlsx');
  
  previousSourceWorkbook = new ExcelDocument();
  await previousSourceWorkbook.loadXLSX(previousSourcefile)

  const latestSourcefile = readFileSync('./test/2. Source File_NEW_CV_Basrah_AB.xlsx');
  latestSourceWorkbook = new ExcelDocument();
  await latestSourceWorkbook.loadXLSX(latestSourcefile)
});

test("Read XLSX worksheet", async () => {
  if(!previousSourceWorkbook) throw new Error('workbook is null');
  previousSourceWorksheet = await previousSourceWorkbook.getWorksheet(1);
  previousSourceRange = previousSourceWorksheet.getRange("B6", "BJ9999999999");

  if(!latestSourceWorkbook) throw new Error('workbook is null');
  latestSourceWorksheet = await latestSourceWorkbook.getWorksheet(1);
  latestSourceRange = latestSourceWorksheet.getRange("B6", "BJ9999999999"); 
});

test("Update XLSX worksheet - Fill up values", async () => {
  const applyRangeFormat = (range: (CellObject | null)[][], reference: (CellObject | null)[]) => {
    for(let rowNo=0; rowNo<range.length; rowNo++) {
      const row = range[rowNo];
      for(let columnNo=0; columnNo<row.length; columnNo++) {
        const style = reference[columnNo]?.style;
  
        if(row[columnNo] === undefined) row[columnNo] = {};

        row[columnNo]!.style = style;
      }
    }
  };

  const file = readFileSync('./test/Comparison of Valve Dimension Table S_DDMMYYYY.xlsx');
  
  templateWorkbook = new ExcelDocument();
  await templateWorkbook.loadXLSX(file);
  

  templateWorksheet7 = await templateWorkbook.getWorksheet(7); 
  const previousWorksheetFormat = templateWorksheet7.getRange('B6', 'BJ6')[0]; 
  applyRangeFormat(previousSourceRange, previousWorksheetFormat); 
  templateWorksheet7.updateRange(previousSourceRange, 'B6'); 

  templateWorksheet6 = await templateWorkbook.getWorksheet(6); 
  const latestWorksheetFormat = templateWorksheet6.getRange('B6', 'BJ6')[0]; 
  applyRangeFormat(latestSourceRange, latestWorksheetFormat); 
  templateWorksheet6.updateRange(latestSourceRange, 'B6');
});

test("Update XLSX worksheet - Fillup formulas", async () => {
  const updateFormula = async (workSheetNo: number): Promise<void> => {
    const templateWorksheet4 = await templateWorkbook.getWorksheet(workSheetNo);
    const template4Formulas = templateWorksheet4.getRange('B6', 'BJ6')[0]; 
    for (let rowNo = 0; rowNo < latestSourceRange.length; rowNo++) {
      const row = template4Formulas;

      for (let columnNo = 0; columnNo < row.length; columnNo++) {
        const column = row[columnNo];
        if(!column?.formula) continue;
        const updatedFormulaValue = column.formula.value.replaceAll('6', (rowNo + 6).toString());

        const cellRef = ExcelColumnConverter.numberToColumn(columnNo + 2) + Number(rowNo + 6);
        templateWorksheet4.updateCell({formula: {value: updatedFormulaValue, type: column.formula.type}, style: column.style}, cellRef);
      }
    }
  };

  await updateFormula(5);
  await updateFormula(4);
  await updateFormula(3);
  await updateFormula(2);

  
  const templateReportWorksheet = await templateWorkbook.getWorksheet(1);
  const templateReportRange = templateReportWorksheet.getRange('A3', 'J3')[0];

  const reportRange: CellObject[][] = [];
  for (let i=0; i< previousSourceRange.length; i++) {
      const previousItem = previousSourceRange[i]; 
      const previousF1 = Number(previousItem[24]?.value);
      const previousWeight = Number(previousItem[48]?.value); 

      const latestItem = latestSourceRange[i];
      const latestF1 = Number(latestItem[24]?.value);
      const latestWeight = Number(latestItem[48]?.value);
      console
      if(
          !isNaN(previousF1) && !isNaN(latestF1) &&
          Math.abs(previousF1 - latestF1) >= 300
      ) {
          const itemNoValue = previousItem[0]?.value ?? '';
          const tagNoValue = previousItem[1]?.value ?? ''; 
          const valveTypeValue = previousItem[4]?.value ?? ''; 
          const partNameValue = previousItem[5]?.value ?? ''; 
          const parameterValue = "F1";
          const descriptionValue = "The height has changed significantly from previous data. Please make sure that the height is correct.";

          const rowCells: CellObject[] = [];
          const heightItems =[itemNoValue, tagNoValue, valveTypeValue, partNameValue, parameterValue, previousF1.toString(), latestF1.toString(), descriptionValue];

          for(let i=0; i<heightItems.length; i++) {
            rowCells.push({value: heightItems[i], style: templateReportRange[i]?.style});
          }

          reportRange.push(rowCells);
      }

      if(
          !isNaN(previousWeight) && !isNaN(latestWeight) &&
          Math.abs(previousWeight - latestWeight) >= 500
      ) {
          const itemNoValue = previousItem[0]?.value ?? '';
          const tagNoValue = previousItem[1]?.value ?? '';
          const valveTypeValue = previousItem[4]?.value ?? '';
          const partNameValue = previousItem[5]?.value ?? '';
          const parameterValue = "Weight";
          const descriptionValue = "The weight has changed significantly from previous data. Please make sure that the weight is correct.";

          const rowCells: CellObject[] = [];
          const weightItems = [itemNoValue, tagNoValue, valveTypeValue, partNameValue, parameterValue, previousWeight.toString(), latestWeight.toString(), descriptionValue]
          
          for(let i=0; i<weightItems.length; i++) {
            rowCells.push({value: weightItems[i], style: templateReportRange[i]?.style});
          }

          reportRange.push(rowCells);
      }
  }
  templateReportWorksheet.updateRange(reportRange, "A3");
  
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