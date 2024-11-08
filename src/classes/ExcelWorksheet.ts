import { CellObject } from "../types";
import ExcelColumnConverter from "./ExcelColumnConverter";
import ExcelSharedStrings from "./ExcelSharedStrings";

export default class ExcelWorksheet {
  private xmlDocument!: Document;
  private worksheetElement!: Element;
  private namespace!: string;
  private sheetDataElement!: Element;
  private sharedStrings!: ExcelSharedStrings;

  public columns!: Array<{min: number, max: number, style: string | null, element: HTMLTableColElement}>;
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

    this.columns = [];
    const columns = this.xmlDocument.querySelector('cols')?.querySelectorAll('col');
    if(columns != undefined) {
      for(let i=0; i<columns?.length; i++) {
        const element = columns[i];
        const min = Number(element.getAttribute('min'));
        const max = Number(element.getAttribute('max'));
        const style = element.getAttribute('style');
        this.columns.push({min, max, style, element});
      }
    }

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

    const cellStyle = cell.style;
    if(cellStyle != undefined) cellElement.setAttribute('s', cellStyle);

    const cellType = cell.type;
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

  public updateRange (cellObjects: (CellObject | null)[][], startCellRef: string = 'A1'): void {
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
    let cellStyle = cellElement.getAttribute('s');
    if(cellStyle === null && this.columns.length > 0) {
      const cellRef = cellElement.getAttribute('r');
      if(!cellRef) throw new Error('Invalid blank cell reference detected');
      
      const {rowIndex, columnIndex} = ExcelColumnConverter.cellRefToIndex(cellRef);
      const column = this.columns.find(({min, max, style, element}) => {
        return columnIndex >= min && columnIndex <= max
      });

      if(column) cellStyle = column.style; //console.log(cellStyle);
    }

    if(!cellValue && !cellFormula && !cellType && !cellStyle) return cellObject;
    
    if(cellValue) cellObject.value = cellValue;
    if(cellFormula) cellObject.formula = cellFormula;
    if(cellStyle) cellObject.style = cellStyle;
    cellObject.type = cellType === 's' ? 'string' : 'number';

    if(cellType === 's') {
      cellObject.value = this.sharedStrings.getIndexString(Number(cellValue));
    }

    return cellObject;
  }

  public getCell (cellRef: string): CellObject | undefined {
      const cell = this.cellsMap.get(cellRef);
      if(!cell) return undefined;

      const cellObject = this.cellElementToObject(cell);
      return cellObject
  }

  public getRange (startCellRef?: string, endCellRef?: string): Array<Array<CellObject | null>> {
    console.log('Retrieving worksheet range values');
    const cellObjectRange: (CellObject | null)[][] = [];

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
/*
    for (let rowNo=0; rowNo<=cellObjectRange.length; rowNo++) {
      const row = cellObjectRange[rowNo];
      if(!row || row.length === 0) continue;

      for (let columnNo=0; columnNo<=row.length; columnNo++) {
        const cellObject = row[columnNo];
        if(cellObject) continue;

        const columnIndex = startColumnIndex + columnNo;

        const column = this.columns.find(({min, max, style, element}) => {
          return columnIndex >= min && columnIndex <= max
        });
  
        if(column) cellObjectRange[rowNo][columnNo] = {type: "string", style: column.style};
      }
    }*/

    return cellObjectRange;
  }

  public toString(): string {
    return new XMLSerializer().serializeToString(this.xmlDocument);
  }
}