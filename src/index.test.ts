import { expect, expectTypeOf, test } from "vitest";
import ExcelDocument from "./index";
import {read, readFileSync} from 'fs';
import ExcelColumnConverter from "./classes/ExcelColumnConverter";
import fs from 'fs';
import ExcelWorksheet from "./classes/ExcelWorksheet";

let workbook: ExcelDocument;
let worksheet: ExcelWorksheet;

test("Load XLSX file", async () => {
  const file = readFileSync('./test/Book1.xlsx');
  
  workbook = new ExcelDocument();
  await workbook.loadXLSX(file)
});

test("Read XLSX worksheet", async () => {
  if(!workbook) throw new Error('workbook is null');
  worksheet = await workbook.getWorksheet(2);
  //const rangeValues = worksheet.getRangeValues();
});

test("Update XLSX worksheet", async () => {
  worksheet.addCell({Value: 30009, Format: {Type: null, Style: null}}, 2, 33);

  //worksheet.addRows([[{Value: 3234, Format: {Type: null, Style: null}}]], 9, 9);
  /*
  for (let i = 1; i < 20; i++) {
    for (let j = 1; j < 20; j++) {
      worksheet.addCell({Value: 696969696969, Format: {Type: null, Style: null}}, i, j);
    }
  }*/
  


});

test("Save XLSX file", async () => {

  const arrayBuffer = await workbook.saveXLSX();
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