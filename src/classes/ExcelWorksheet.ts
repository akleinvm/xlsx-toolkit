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

  public fromXML(xmlString: string) {
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

  public updateCell (cell: CellObject, cellRef: string): void {
    const {rowIndex: rowNo, columnIndex: columnNo} = ExcelColumnConverter.cellRefToIndex(cellRef);

    let rowElement = this.rowsMap.get(rowNo);
    if(!rowElement) {
      rowElement = this.xmlDocument.createElementNS(this.namespace, 'row');
      rowElement.setAttribute('r', rowNo.toString());
      rowElement.setAttribute('spans', `${1}:${columnNo}`);

      const nextSibling = Array.from(this.rowsMap).find(([key]) => key > rowNo)?.[1];
      
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

    const cellStyle = cell.format?.style;
    if(cellStyle) cellElement.setAttribute('s', cellStyle);

    const cellType = cell.format?.type;
    if(cellType === 'string') cellElement.setAttribute('t', 's');
      
    const cellChildren: Element[] = [];
    
    if(cell.formula) {
      const formula = cell.formula;
      const formulaElement = this.xmlDocument.createElementNS(this.namespace, 'f');
      formulaElement.textContent = formula;
      cellChildren.push(formulaElement);
    } else if (cell.value) {
      let value = cell.value;
      if(cellType === 'string') value = this.sharedStrings.getStringIndex(cell.value).toString();
      const valueElement = this.xmlDocument.createElementNS(this.namespace, 'v');
      valueElement.textContent = value;
      cellChildren.push(valueElement);
    }
      

    cellElement.replaceChildren(...cellChildren);
    rowElement.appendChild(cellElement);
    this.cellsMap.set(cellReference, cellElement);
  }

  public updateRange (cellObjects: (CellObject | undefined)[][], startCellRef: string = 'A1'): void {
    const {rowIndex, columnIndex} = ExcelColumnConverter.cellRefToIndex(startCellRef);

    for (let rowNo = 0; rowNo < cellObjects.length; rowNo++) {
      const row = cellObjects[rowNo];
      if(!row?.length) continue;
      for (let columnNo = 0; columnNo < row.length; columnNo++) {
        const cell = row[columnNo];
        if(!cell) continue;

        const cellRef = ExcelColumnConverter.numberToColumn(columnIndex + columnNo) + Number(rowIndex + rowNo);
        this.updateCell(cell, cellRef);
      }
    }

  }

  public cellElementToObject (cellElement: Element): CellObject {
    const cellObject: CellObject = {}

    if(!cellElement) return cellObject;

    const cellValue = cellElement.querySelector('v')?.textContent;
    const cellFormula = cellElement.querySelector('f')?.textContent;
    const cellType = cellElement.getAttribute('t');
    const cellStyle = cellElement.getAttribute('s');

    if(!cellValue && !cellFormula && !cellType && !cellStyle) return cellObject;
    
    if(cellValue) cellObject.value = cellValue;
    if(cellFormula) cellObject.value = cellFormula;
    
    if(cellType) cellObject.format = {type: cellType === 's' ? 'string' : 'number', style: cellStyle};
    if(cellType === 's') {
      cellObject.value = this.sharedStrings.getIndexString(Number(cellValue));
      cellObject.formula = cellFormula!;
      cellObject.format = {type: 'string', style: cellStyle}
    } else {
      cellObject.formula = cellFormula!;
      cellObject.format = {type: 'number', style: cellStyle};
    }

    return cellObject;
  }

  public getCell (cellRef: string): CellObject | undefined {
      const cell = this.cellsMap.get(cellRef);
      if(!cell) return undefined;

      const cellObject = this.cellElementToObject(cell);
      return cellObject
  }

  public getRange (startCellRef?: string, endCellRef?: string): Array<Array<CellObject | undefined>> {
    console.log('Retrieving worksheet range values');
    const cellObjectRange: CellObject[][] = [];

    const {rowIndex: startRowIndex, columnIndex: startColumnIndex} = 
      startCellRef ? ExcelColumnConverter.cellRefToIndex(startCellRef) :
      {rowIndex: 1, columnIndex: 1}
    ;

    const {rowIndex: endRowIndex, columnIndex: endColumnIndex} = 
      endCellRef ? ExcelColumnConverter.cellRefToIndex(endCellRef) : 
      {rowIndex: 999999999, columnIndex: 999999999}
    ;

    for(const [cellRef, cellElement] of this.cellsMap) {
      const {rowIndex, columnIndex} = ExcelColumnConverter.cellRefToIndex(cellRef);

      if (
        rowIndex < startRowIndex ||
        rowIndex > endRowIndex && 
        columnIndex < startColumnIndex && 
        columnIndex > endColumnIndex
      ) continue;

      const rowNo = rowIndex - startRowIndex;
      const columnNo = columnIndex - startColumnIndex;

      const cellObject = this.cellElementToObject(cellElement);
      if(!cellObjectRange[rowNo]) cellObjectRange[rowNo] = [];
      cellObjectRange[rowNo][columnNo] = cellObject
    }
    return cellObjectRange;
  }

  public toString(): string {
    return new XMLSerializer().serializeToString(this.xmlDocument);
  }
}