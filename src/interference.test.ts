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
  valveType: string, 
  parameter: string, 
  assemblyType: string, 
  connectionType: string, 
  inlet: number, 
  outlet: number, 
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
      const style = reference[columnNo]?.style;

      if(row[columnNo] === undefined) row[columnNo] = {};

      row[columnNo]!.style = style;
    }
  }
};

test("Load XLSX file", async () => {
  const sourcefile = readFileSync('./test/1. Source File_OLD_CV_Basrah_2.0.xlsx');
  sourceWorkbook = new ExcelDocument();
  await sourceWorkbook.loadXLSX(sourcefile);

  const templatefile = readFileSync('./test/Interference report template.xlsx');
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
    const [idCode, valveType, parameter, assemblyType, connectionType, inlet, outlet] = row.map((content) => content?.value);
    const dimension: {[key: string]: number} = {};
    for(let columnNo=8; columnNo<14; columnNo++) {
      const columnName = ffRange[0][columnNo]?.value?.toString().replaceAll('#','');
      if(!columnName) throw new Error('Invalid FF dimension name');
      dimension[columnName] = Number(row[columnNo]?.value)
    }
    ffDimensionTable.push({
      valveType: valveType === undefined ? '' :  valveType.toString(),
      parameter: parameter === undefined ? '' :  parameter.toString(),
      assemblyType: assemblyType === undefined ? '' :  assemblyType.toString(),
      connectionType: connectionType === undefined ? '' :  connectionType.toString(), 
      inlet: Number(inlet), 
      outlet: Number(outlet), 
      dimensions: dimension
    });
  }

  const flangeWorksheet = await templateWorkbook.getWorksheet(3);
  const flangeRange = flangeWorksheet.getRange('A2', 'N999999999');
  for(let rowNo=2; rowNo<flangeRange.length; rowNo++) {
    const row = flangeRange[rowNo];
    if(!row) continue;
    const [nps, dn] = row.map((content) => Number(content?.value))
    const rating: {[key: string]: number} = {};
    for(let columnNo=2; columnNo<8; columnNo++) {
      const columnName = flangeRange[1][columnNo]?.value;
      if(!columnName) throw new Error('Invalid flange rating name');
      rating[columnName.toString()] = Number(row[columnNo]?.value)
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

      const A = Number(row[17]?.value) || 0; 
      const B = Number(row[18]?.value) || 0; 
      const L = Number(row[21]?.value) || 0; 
      const F1 = Number(row[24]?.value) || 0; 
      const F2 = Number(row[25]?.value) || 0; 
      const G = Number(row[26]?.value) || 0; 
      const PS2 = Number(row[34]?.value) || 0; 
      const PS4 = Number(row[36]?.value) || 0; 
      const HW2 = Number(row[46]?.value) || 0; 
      const HW3 = Number(row[47]?.value) || 0; 

      const connectionType = row[11]?.value; 
      const inletSize = Number(row[8]?.value); 
      const outletSize = Number(row[9]?.value); 
      const measuringStandard = row[10]?.value === 'in' ? 'nps' : row[10]?.value === 'mm' ? 'dn' : '';
      const rating = Number(row[7]?.value?.toString().replaceAll('#', '')); 

      let ffDimension: number | undefined;
      if(connectionType && !isNaN(inletSize) && !isNaN(outletSize)) ffDimension = ffDimensionTable.find((dimension) => {
        return dimension.connectionType === connectionType &&
        dimension.valveType === partName &&
        dimension.inlet === inletSize &&
        dimension.outlet === outletSize
      })?.dimensions[rating];

      let flangeRating: number | undefined;
      if(!isNaN(outletSize) && measuringStandard != '') flangeRating = flangeTable.find((flange) => {
        return flange[measuringStandard] === outletSize;
      })?.rating[rating];

      let pipeOuterDiameter: number | undefined;
      if(!isNaN(outletSize) && measuringStandard != '') pipeOuterDiameter = pipeTable.find((pipe) => {
        return pipe[measuringStandard] === outletSize;
      })?.outerDiameterMM;

      //console.log(measuringStandard);
      //console.log(`item:${item} tagNo:${tagNo} valveType:${valveType} partName:${partName} A:${A} B:${B} L:${L} F1:${F1} F2:${F2} G:${G} PS2:${PS2} PS4:${PS4} HW2:${HW2} HW3:${HW3} connectionType:${connectionType} inletSize:${inletSize} outletSize:${outletSize} rating:${rating}`)
      //console.log(`ffDimension:${ffDimension} flangeRating:${flangeRating} pipeOuterDiameter:${pipeOuterDiameter}`);
      //console.log(JSON.stringify(flangeTable.slice(0,3)));

      const checkArray: {checkItemNo: number, description: string, parameter: string, condition: boolean}[] = [
        {
          checkItemNo: 1,
          description: "F-to-F dimension must be rechecked.",
          parameter: "A",
          condition: !!ffDimension && (!isNaN(A) && !isNaN(ffDimension) && !(Math.abs(A - ffDimension) < 1))
        },
        {
          checkItemNo: 2,
          description: "Positioner height must be rechecked.",
          parameter: "PS2",
          condition: !!valveType && valveType != 'IV' && (
            ['I', 'II', 'III'].includes(valveType.toString()) && !(Math.abs(F1*0.55-PS2) < 250 && PS2-PS4 > 250) || 
            valveType === 'V' && !(Math.abs(F2-(PS2-PS4)/2) < 250)
          )
        },
        {
          checkItemNo: 3,
          description: 'HW height must be rechecked.',
          parameter: "HW2",
          condition: !!valveType && valveType != 'IV' && (
            ['I', 'II', 'III'].includes(valveType.toString()) && !(Math.abs(F1*0.55-HW2) < 250 && HW2-HW3/2 > 250) || 
            valveType === 'V' && !(Math.abs(F2-HW2) < 250)
        )
        },
        {
          checkItemNo: 4,
          description: "There is interference or very little clearance between the actuator and process piping.",
          parameter: "G (or F2)",
          condition: !!valveType && !!partName && valveType != 'IV' && !!flangeRating && !isNaN(flangeRating) && 
          ['Ball', 'Butterfly'].includes(partName.toString()) && !(F2-G/2-flangeRating/2 > 100)
        },
        {
          checkItemNo: 5,
          description: "There is interference or very little clearance between the HW and process piping",
          parameter: "HW3 (or F2)",
          condition: !!valveType && !!partName && valveType != 'IV' && !!pipeOuterDiameter && !isNaN(pipeOuterDiameter) && 
          ['Ball', 'Butterfly'].includes(partName.toString()) && !(F2-HW3/2-pipeOuterDiameter/2 > 100)
        },
        {
          checkItemNo: 6,
          description: "Support legs too short.",
          parameter: "L (or B)",
          condition: !!valveType && partName === 'Ball' && valveType != 'IV' && !(B-L/2 > 0)
        }
      ]
      
      for(const {condition, description, checkItemNo, parameter} of checkArray) {
        if(condition) {
          const rowCells: CellObject[] = [];
          const rowItems =[item, tagNo, valveType, partName, checkItemNo, parameter, description];

          for(let i=0; i<rowItems.length; i++) {
            rowCells.push({value: rowItems[i]});
          }

          reportRange.push(rowCells);
        }
      }
  }
  const reportWorksheetFormat = reportWorksheet.getRange('A3', 'I3')[0]; 
  applyRangeFormat(reportRange, reportWorksheetFormat); 
  reportWorksheet.updateRange(reportRange, "A3");
  
});

test("Save XLSX file", async () => {

  const arrayBuffer = await templateWorkbook.saveXLSX();
  if(!arrayBuffer) return;
  const buffer = Buffer.from(arrayBuffer);

  fs.writeFile('./test/output/interferencetest.xlsx', buffer, (error) => {
    if (error) {
      console.error('Error saving the file:', error);
    } else {
        console.log('File saved successfully!');
    }
  })
});