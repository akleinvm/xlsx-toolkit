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

let previousWorkbook: ExcelDocument;
let previousWorksheet: ExcelWorksheet;
let previousRange: (CellObject | null)[][];

let latestWorkbook: ExcelDocument;
let latestWorksheet: ExcelWorksheet;
let latestRange: (CellObject | null)[][];

let templateWorkbook: ExcelDocument;
let templateNewWorksheet: ExcelWorksheet;
let templateOldWorksheet: ExcelWorksheet;

test("Load XLSX file", async () => {
  const latestfile = readFileSync('./test/1. Source File_OLD_CV_Basrah_2.0.xlsx');
  const previousfile = readFileSync('./test/2. Source File_NEW_CV_Basrah_AB.xlsx');
  
  previousWorkbook = new ExcelDocument();
  await previousWorkbook.loadXLSX(previousfile)

  latestWorkbook = new ExcelDocument();
  await latestWorkbook.loadXLSX(latestfile)
});

test("Read XLSX worksheet", async () => {
  if(!previousWorkbook) throw new Error('workbook is null');
  previousWorksheet = await previousWorkbook.getWorksheet(1);
  previousRange = previousWorksheet.getRange("B6", "BJ9999999999");

  if(!latestWorkbook) throw new Error('workbook is null');
  latestWorksheet = await latestWorkbook.getWorksheet(1);
  latestRange = latestWorksheet.getRange("B6", "BJ9999999999"); 
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

  const file = readFileSync('./test/Comparison report template.xlsx');
  
  templateWorkbook = new ExcelDocument();
  await templateWorkbook.loadXLSX(file);
  

  templateOldWorksheet = await templateWorkbook.getWorksheet(7); 
  const previousWorksheetFormat = templateOldWorksheet.getRange('B6', 'BJ6')[0]; 
  applyRangeFormat(previousRange, previousWorksheetFormat); 
  templateOldWorksheet.updateRange(previousRange, 'B6'); 

  templateNewWorksheet = await templateWorkbook.getWorksheet(6); 
  const latestWorksheetFormat = templateNewWorksheet.getRange('B6', 'BJ6')[0]; 
  applyRangeFormat(latestRange, latestWorksheetFormat); 
  templateNewWorksheet.updateRange(latestRange, 'B6');
}, 10000);

test("Update XLSX worksheet - Fillup formulas", async () => {
  const updateFormula = async (workSheetNo: number): Promise<void> => {
    const templateWorksheet4 = await templateWorkbook.getWorksheet(workSheetNo);
    const template4Formulas = templateWorksheet4.getRange('B6', 'BJ6')[0]; 
    for (let rowNo = 0; rowNo < latestRange.length; rowNo++) {
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
  for (let i=0; i< previousRange.length; i++) {
      const previousItem = previousRange[i]; 
      const previousF1 = Number(previousItem[24]?.value);
      const previousWeight = Number(previousItem[48]?.value); 

      const latestItem = latestRange[i];
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
  
}, 10000);

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