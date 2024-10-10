import { CellObject } from "../types";
import ExcelColumnConverter from "./ExcelColumnConverter";
import ExcelSharedStrings from "./ExcelSharedStrings";

export default class ExcelWorksheet {
  private xmlDocument!: Document;
  private worksheetElement!: Element;
  private namespace!: string;
  private sheetDataElement!: Element;
  private sharedStrings!: ExcelSharedStrings;

  public rowsMap!: Map<number, Element>;
  public cellsMap!: Map<string, Element>;

  constructor(sharedStrings: ExcelSharedStrings) {
    this.sharedStrings = sharedStrings;
  }

  public async fromXML(xmlString: string) {
    this.xmlDocument = new DOMParser().parseFromString(xmlString, "text/xml");

    this.worksheetElement = this.xmlDocument.getElementsByTagName('worksheet')[0];
    this.namespace = this.worksheetElement.getAttribute('xmlns') ?? "";

    this.sheetDataElement = this.xmlDocument.getElementsByTagName("sheetData")[0];
    this.rowsMap = new Map();
    const rows = this.sheetDataElement.querySelectorAll('row');

    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      const rowNo = Number(row.getAttribute('r'));
      this.rowsMap.set(rowNo, row);
    }

    this.cellsMap = new Map();
    const cells = this.sheetDataElement.querySelectorAll('c');

    for (let i = 0; i < cells.length; i++) {
      const cell = cells[i];
      const cellRef = cell.getAttribute('r');
      
      if(cellRef) this.cellsMap.set(cellRef, cell);
    }
  }

  public updateCellValue (cell: CellObject, rowNo: number, columnNo: number): void {
    let rowElement = this.rowsMap.get(rowNo);
    if(!rowElement) {
      rowElement = this.xmlDocument.createElementNS(this.namespace, 'row');
      rowElement.setAttribute('r', rowNo.toString());
      rowElement.setAttribute('spans', `${1}:${columnNo}`);

      const nextSibling = Array.from(this.rowsMap).find(([key, element]) => key > rowNo)?.[1];
      
      if(!nextSibling) {
        this.sheetDataElement.appendChild(rowElement);
      } else {
        this.sheetDataElement.insertBefore(rowElement, nextSibling);
      }
      
      this.rowsMap.set(rowNo, rowElement);
    } else {
      const [minSpan, maxSpan] = rowElement.getAttribute('spans')?.split(':') ?? [columnNo, columnNo];
      const spans = `${Math.min(columnNo, Number(minSpan)).toString()}:${Math.max(columnNo, Number(maxSpan))}`;
      rowElement.setAttribute('spans', spans);
    }

    const cellReference = ExcelColumnConverter.numberToColumn(columnNo) + Number(rowNo);
    let cellElement = this.cellsMap.get(cellReference);
    if(!cellElement) {
      cellElement = this.xmlDocument.createElementNS(this.namespace, 'c');
      cellElement.setAttribute('r', cellReference);
    }

    const cellStyle = cell.Format.Style;
    if(!cellStyle) cellElement.removeAttribute('s'); 
    else cellElement.setAttribute('s', cellStyle);

    const cellType = cell.Format.Type;
    if(!cellType) cellElement.removeAttribute('t'); 
    else cellElement.setAttribute('t', cellType);
        
    let value = cell.Value;
    if(cellType === 's') value = this.sharedStrings.getStringIndex(cell.Value).toString();

    const valueElement = this.xmlDocument.createElementNS(this.namespace, 'v');
    valueElement.textContent = value;
    cellElement.replaceChildren(valueElement);
    rowElement.appendChild(cellElement);
    this.cellsMap.set(cellReference, cellElement);
  }




  public getRangeValues (): string[][] {
    console.log('Retrieving worksheet range values');

    const output: string[][] = [];

    for(const [key, cell] of this.cellsMap) {
      const {RowIndex, ColumnIndex} = ExcelColumnConverter.cellRefToIndex(key);
      const rowNo = RowIndex - 1;
      const columnNo = ColumnIndex - 1;

      const valueElement = cell.querySelector('v');
      if(!valueElement?.textContent) continue;

      let cellValue = valueElement.textContent;

      const cellType = cell.getAttribute('t');
      if(cellType === 's') cellValue = this.sharedStrings.getIndexString(Number(cellValue));

      if(!output[rowNo]) output[rowNo] = [];
      output[rowNo][columnNo] = cellValue;
    }
    return output;
  }


  public toString(): string {
    return new XMLSerializer().serializeToString(this.xmlDocument);
  }
}