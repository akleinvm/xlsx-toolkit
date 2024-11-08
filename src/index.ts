import ExcelDocument from "./classes/ExcelDocument";
import ExcelColumnConverter from "./classes/ExcelColumnConverter";
import { CellObject, CellIndex } from "./types";

export {ExcelDocument, ExcelColumnConverter, CellObject, CellIndex};

export async function parseXLSX (arrayBuffer: ArrayBuffer): Promise<ExcelDocument> {
    const excelDocument = new ExcelDocument();
    await excelDocument.loadXLSX(arrayBuffer);

    return excelDocument;
}