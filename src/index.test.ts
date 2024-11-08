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

let sourceWorkbook: ExcelDocument;
let sourceWorksheet: ExcelWorksheet;
let sourceRange: (CellObject | null)[][];

let templateWorkbook: ExcelDocument;
let reportWorksheet: ExcelWorksheet;
const ffDimensionTable: {
  idCode: string, 
  valveType: string, 
  parameter: string, 
  assemblyType: string, 
  connectionType: string, 
  inlet: number, 
  outlet: number, 
  connectionSize: string,
  dimensions: {[key: string]: number}
}[] = [];
const flangeTable: {
  nps: number,
  dn: number,
  rating: {[key: string]: number}
}[] = [];
const pipeTable: {
  nps: number,
  dn: number,
  outerDiameterIN: number, 
  outerDiameterMM: number
}[] = [];

const applyRangeFormat = (range: (CellObject | null)[][], reference: (CellObject | null)[]) => {
  for(let rowNo=0; rowNo<range.length; rowNo++) {
    const row = range[rowNo];
    for(let columnNo=0; columnNo<row.length; columnNo++) {
      const type = reference[columnNo]?.type;
      const style = reference[columnNo]?.style;

      if(row[columnNo] === undefined) row[columnNo] = {};

      row[columnNo]!.type = type;
      row[columnNo]!.style = style;
    }
  }
};

test("Load XLSX file", async () => {
  const sourcefile = readFileSync('./test/1. Source File_OLD_CV_Basrah_2.0.xlsx');
  sourceWorkbook = new ExcelDocument();
  await sourceWorkbook.loadXLSX(sourcefile);

  const templatefile = readFileSync('./test/Interference Check R1a.xlsx');
  templateWorkbook = new ExcelDocument();
  await templateWorkbook.loadXLSX(templatefile);
  
});

test("Read XLSX worksheet", async () => {
  if(!sourceWorkbook) throw new Error('workbook is null');
  sourceWorksheet = await sourceWorkbook.getWorksheet(1);
  sourceRange = sourceWorksheet.getRange("B6", "BJ9999999999"); 

  if(!templateWorkbook) throw new Error('workbook is null');
  reportWorksheet = await templateWorkbook.getWorksheet(5);

  const ffWorksheet = await templateWorkbook.getWorksheet(4);
  const ffRange = ffWorksheet.getRange('A3', 'N999999999');
  for(let rowNo=1; rowNo<ffRange.length; rowNo++) {
    const row = ffRange[rowNo];
    if(!row) continue;
    const [idCode, valveType, parameter, assemblyType, connectionType, inlet, outlet, connectionSize] = row.map((content) => content?.value ?? '')
    const dimension = {};
    for(let columnNo=8; columnNo<14; columnNo++) {
      const columnName = ffRange[0][columnNo]?.value?.replaceAll('#','');
      if(!columnName) throw new Error('Invalid FF dimension name');
      dimension[columnName] = Number(row[columnNo]?.value)
    }
    ffDimensionTable.push({
      idCode: idCode, 
      valveType: valveType, 
      parameter: parameter, 
      assemblyType: assemblyType, 
      connectionType: connectionType, 
      inlet: Number(inlet), 
      outlet: Number(outlet), 
      connectionSize: connectionSize,
      dimensions: dimension
    });
  }

  const flangeWorksheet = await templateWorkbook.getWorksheet(3);
  const flangeRange = flangeWorksheet.getRange('A2', 'N999999999');
  for(let rowNo=2; rowNo<flangeRange.length; rowNo++) {
    const row = flangeRange[rowNo];
    if(!row) continue;
    const [nps, dn] = row.map((content) => Number(content?.value))
    const rating = {};
    for(let columnNo=2; columnNo<8; columnNo++) {
      const columnName = flangeRange[1][columnNo]?.value;
      if(!columnName) throw new Error('Invalid flange rating name');
      rating[columnName] = Number(row[columnNo]?.value)
    }
    flangeTable.push({
      nps: nps, 
      dn: dn, 
      rating: rating
    });
  }

  const pipeWorksheet = await templateWorkbook.getWorksheet(2);
  const pipeRange = pipeWorksheet.getRange('A2', 'N999999999');
  for(let rowNo=2; rowNo<pipeRange.length; rowNo++) {
    const row = pipeRange[rowNo];
    if(!row) continue;
    const [nps, dn, outerDiameterIN, outerDiameterMM] = row.map((content) => Number(content?.value))
    pipeTable.push({
      nps: nps, 
      dn: dn, 
      outerDiameterIN: Number(outerDiameterIN),
      outerDiameterMM: Number(outerDiameterMM)
    });
  }
  //console.log(JSON.stringify(pipeTable));

});

test("Update XLSX worksheet - Fill up values", async () => {
  const applyRangeFormat = (range: (CellObject | null)[][], reference: (CellObject | null)[]) => {
    for(let rowNo=0; rowNo<range.length; rowNo++) {
      const row = range[rowNo];
      for(let columnNo=0; columnNo<row.length; columnNo++) {
        const type = reference[columnNo]?.type;
        const style = reference[columnNo]?.style;
  
        if(row[columnNo] === undefined) row[columnNo] = {};

        row[columnNo]!.type = type;
        row[columnNo]!.style = style;
      }
    }
  };  

  const dimensionWorksheet = await templateWorkbook.getWorksheet(1);
  const dimensionWorksheetFormat = dimensionWorksheet.getRange('B6', 'BJ6')[0]; 
  applyRangeFormat(sourceRange, dimensionWorksheetFormat); //console.log(JSON.stringify(sourceRange));
  dimensionWorksheet.updateRange(sourceRange, 'B6');
});



test("Update XLSX worksheet - Fill up report", async () => {
  const reportRange: CellObject[][] = [];
  for (let rowNo=0; rowNo<sourceRange.length; rowNo++) {
      const row = sourceRange[rowNo];
      if(!row) continue;
      
      const item = row[0]?.value;
      const tagNo = row[1]?.value;
      const valveType = row[4]?.value;
      const partName = row[5]?.value;

      const A = Number(row[17]?.value); 
      const B = Number(row[18]?.value);
      const L = Number(row[21]?.value);
      const F1 = Number(row[24]?.value);
      const F2 = Number(row[25]?.value);
      const G = Number(row[26]?.value);
      const PS2 = Number(row[34]?.value);
      const PS4 = Number(row[36]?.value);
      const HW2 = Number(row[46]?.value);
      const HW3 = Number(row[47]?.value);

      const connectionType = row[11]?.value; 
      const inletSize = Number(row[8]?.value); 
      const outletSize = Number(row[9]?.value); 
      const rating = Number(row[7]?.value?.replaceAll('#', '')); 

      let ffDimension: number | undefined;
      if(connectionType && !isNaN(inletSize) && !isNaN(outletSize)) ffDimension = ffDimensionTable.find((dimension) => {
        return dimension.connectionType === connectionType &&
        dimension.inlet === inletSize &&
        dimension.outlet === outletSize
      })?.dimensions[rating];

      let flangeRating: number | undefined;
      if(!isNaN(outletSize)) flangeRating = flangeTable.find((flange) => {
        return flange.nps === outletSize;
      })?.rating[rating];

      let pipeOuterDiameter: number | undefined;
      if(!isNaN(outletSize)) pipeOuterDiameter = pipeTable.find((pipe) => {
        return pipe.nps === outletSize;
      })?.outerDiameterMM;

      const checkArray: {checkItemNo: string, description: string, parameter: string, condition: boolean}[] = [
        {
          checkItemNo: '1',
          description: "F-to-F dimension must be rechecked.",
          parameter: "A",
          condition: !!ffDimension && (!isNaN(A) && !isNaN(ffDimension) && Math.abs(A - ffDimension) < 1)
        },
        {
          checkItemNo: '2',
          description: "Positioner height must be rechecked.",
          parameter: "PS2",
          condition: !!valveType && !isNaN(F1) && !isNaN(F2) && !isNaN(PS2) && !isNaN(PS4) && (
            ['I', 'II', 'III'].includes(valveType) && Math.abs(F1*0.55-PS2) < 250 && PS2-PS4 > 250 || 
            valveType === 'V' && Math.abs(F2-(PS2-PS4)/2) < 250
          )
        },
        {
          checkItemNo: '3',
          description: 'HW height must be rechecked.',
          parameter: "HW2",
          condition: !!valveType && !isNaN(F1) && !isNaN(F2) && !isNaN(HW2) && !isNaN(HW3) && (
            ['I', 'II', 'III'].includes(valveType) && Math.abs(F1*0.55-HW2) < 250 && HW2-HW3/2 > 250 || 
            valveType === 'V' && Math.abs(F2-HW2) < 250
        )
        },
        {
          checkItemNo: '4',
          description: "There is interference or very little clearance between the actuator and process piping.",
          parameter: "G (or F2)",
          condition: !!valveType && !isNaN(F2) && !isNaN(G) && !!flangeRating && !isNaN(flangeRating) && 
          ['Ball', 'Butterfly'].includes(valveType) && F2-G/2-flangeRating/2 >100
        },
        {
          checkItemNo: '5',
          description: "There is interference or very little clearance between the HW and process piping",
          parameter: "HW3 (or F2)",
          condition: !!valveType && !isNaN(F2) && !isNaN(HW3) && !!pipeOuterDiameter && !isNaN(pipeOuterDiameter) && 
          ['Ball', 'Butterfly'].includes(valveType) && F2-HW3/2-pipeOuterDiameter/2 > 100
        },
        {
          checkItemNo: '6',
          description: "Support legs too short.",
          parameter: "L (or B)",
          condition: !!valveType && valveType === 'Ball' && !isNaN(B) && !isNaN(L) && B-L/2>0
        }
      ]
      
      for(const {condition, description, checkItemNo, parameter} of checkArray) {
        if(condition) {
          const rowCells: CellObject[] = [];
          const rowItems =[item, tagNo, valveType, partName, checkItemNo, parameter, description];

          for(let i=0; i<rowItems.length; i++) {
            rowCells.push({value: rowItems[i], type: 'string'});
          }

          reportRange.push(rowCells);
        }
      }
  }
  const previousWorksheetFormat = reportWorksheet.getRange('A3', 'I3')[0]; 
  applyRangeFormat(reportRange, previousWorksheetFormat); 
  reportWorksheet.updateRange(reportRange, "A3");
  
});
/*
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
        const type = reference[columnNo]?.type;
        const style = reference[columnNo]?.style;
  
        if(row[columnNo] === undefined) row[columnNo] = {};

        row[columnNo]!.type = type;
        row[columnNo]!.style = style;
      }
    }
  };

  const file = readFileSync('./test/Comparison of Valve Dimension Table S_DDMMYYYY.xlsx');
  
  templateWorkbook = new ExcelDocument();
  await templateWorkbook.loadXLSX(file);
  

  templateWorksheet7 = await templateWorkbook.getWorksheet(7);
  const previousWorksheetFormat = templateWorksheet7.getRange('B6', 'BJ6')[0]; //console.log(JSON.stringify(previousWorksheetFormat));
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
        const updatedColumn = column.formula.replaceAll('6', (rowNo + 6).toString());

        const cellRef = ExcelColumnConverter.numberToColumn(columnNo + 2) + Number(rowNo + 6);
        templateWorksheet4.updateCell({formula: updatedColumn, type: column.type, style: column.style}, cellRef);
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
            rowCells.push({value: heightItems[i], type: templateReportRange[i]?.type, style: templateReportRange[i]?.style});
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
            rowCells.push({value: weightItems[i], type: templateReportRange[i]?.type, style: templateReportRange[i]?.style});
          }

          reportRange.push(rowCells);
      }
  }
  templateReportWorksheet.updateRange(reportRange, "A3");
  
});
*/
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