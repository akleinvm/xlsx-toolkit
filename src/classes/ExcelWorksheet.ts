import { CellObject } from "../types";
import ExcelColumnConverter from "./ExcelColumnConverter";
import ExcelSharedStrings from "./ExcelSharedStrings";

export default class ExcelWorksheet {
  private xmlDocument!: Document;
  private worksheetElement!: Element;
  private namespace!: string;
  private sheetDataElement!: Element;
  private sharedStrings!: ExcelSharedStrings;

  private rowsMap!: Map<number, Element>;
  private cellsMap!: Map<string, Element>;

  constructor(sharedStrings: ExcelSharedStrings) {
    this.sharedStrings = sharedStrings;
  }

  public fromXML(xmlString: string) {
    this.xmlDocument = new DOMParser().parseFromString(xmlString, "text/xml");

    this.worksheetElement = this.xmlDocument.getElementsByTagName('worksheet')[0];
    this.namespace = this.worksheetElement.getAttribute('xmlns') ?? "";

    this.sheetDataElement = this.xmlDocument.getElementsByTagName("sheetData")[0];
    this.rowsMap = new Map();
    const rows = this.sheetDataElement.getElementsByTagName('row');
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i];
      this.rowsMap.set(Number(row.getAttribute('r')), row);
    }

    this.cellsMap = new Map();
    const cells = this.sheetDataElement.getElementsByTagName('c');
    for (let i = 0; i < cells.length; i++) {
      const cell = cells[i];
      this.cellsMap.set(cell.getAttribute('r') ?? '', cell);
    }
  }

  public addCell (cell: CellObject, rowNo: number, columnNo: number): void {
    let rowElement = this.rowsMap.get(rowNo);
    if(!rowElement) {
      rowElement = this.xmlDocument.createElementNS(this.namespace, 'row');
      rowElement.setAttribute('r', rowNo.toString());
      rowElement.setAttribute('spans', `${1}:${columnNo}`);
      this.sheetDataElement.appendChild(rowElement);
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
        
    const valueElement = this.xmlDocument.createElementNS(this.namespace, 'v');
    valueElement.textContent = cell.Value.toString();
    cellElement.replaceChildren(valueElement);
    console.log(cellElement.textContent);
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

/*

  private getValues(rangeStart: string, rangeEnd: string): Array<Array<string>> {
    const startIndex = ExcelColumnConverter.cellRefToIndex(rangeStart);
    const minRowNo = startIndex.RowIndex; 
    const minColumnNo = startIndex.ColumnIndex; 

    const endIndex = ExcelColumnConverter.cellRefToIndex(rangeEnd);
    const maxRowNo = endIndex.RowIndex; 
    const maxColumnNo = endIndex.ColumnIndex; 

    const output = new Array<Array<string>>();
    for (let rowNo = minRowNo; rowNo <= maxRowNo; rowNo++) {
      const arrayRowNo = rowNo - minRowNo;
      output[arrayRowNo] = [];

      for (let columnNo = minColumnNo; columnNo <= maxColumnNo; columnNo++) {
        const currentColumnRef = ExcelColumnConverter.numberToColumn(columnNo);
        const currentCell = this.cellsMap.get(currentColumnRef + rowNo.toString());
        
        let cellValue = '';
        const cell = currentCell?.getElementsByTagName('v')[0];
        if(cell) {
          cellValue = cell.textContent ?? '';
          const cellType = currentCell?.getAttribute('t');
          if(cellType === 's') cellValue = this.sharedStrings.getIndexString(Number(cellValue));
        }
        
        const arrayColumnNo = columnNo - minColumnNo;
        output[arrayRowNo][arrayColumnNo] = cellValue;
      }
    }
    
    return output;
  }
    */

  /*
  
  public addRows(rows: Array<Array<CellObject>>, rowStartIndex: number, columnStartIndex: number): void {
    for(let i = 0; i < rows.length; i++) {
      const rowIndex = rowStartIndex + i;
      const row = rows[i];

      let rowElement = this.rowsMap.get(rowIndex);
      if(!rowElement) {
        rowElement = this.xmlDocument.createElementNS(this.namespace, 'row');
        rowElement.setAttribute('r', rowIndex.toString());
        rowElement.setAttribute('spans', `${columnStartIndex}:${columnStartIndex + row.length}`);
        this.sheetDataElement.appendChild(rowElement);
        this.rowsMap.set(rowIndex, rowElement);
      } else {
        const [minSpan, maxSpan] = rowElement.getAttribute('spans')?.split(':') ?? [];
        const spans = `${Math.min(columnStartIndex, Number(minSpan)).toString()}:${Math.max(columnStartIndex + row.length - 1, Number(maxSpan))}`;
        rowElement.setAttribute('spans', spans);
      }
      
      for(let j = 0; j < row.length; j++) {
        const columnIndex = columnStartIndex + j;
        
        const cellReference = ExcelColumnConverter.numberToColumn(columnIndex) + rowIndex;
        let cellElement = this.cellsMap.get(cellReference);
        if(!cellElement) {
          cellElement = this.xmlDocument.createElementNS(this.namespace, 'c');
          cellElement.setAttribute('r', cellReference);
        }
        
        const cellStyle = row[j].Format.Style;
        if(cellStyle) cellElement.setAttribute('s', cellStyle ?? "");

        const cellType = row[j].Format.Type;
        if(cellType) cellElement.setAttribute('t', cellType ?? "");
        
        const valueElement = this.xmlDocument.createElementNS(this.namespace, 'v');
        valueElement.textContent = row[j].Value.toString();
        cellElement.replaceChildren(valueElement);
        rowElement.appendChild(cellElement);
        
        this.cellsMap.set(cellReference, cellElement);
      }
    }
  }*/