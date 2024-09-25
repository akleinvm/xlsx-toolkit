import { CellObject } from "../types";
import { ExcelColumnConverter } from "./ExcelColumnConverter";

export class ExcelWorksheet {
  private xmlDoc!: Document;
  private worksheetElement!: Element;
  private namespace!: string;
  private sheetDataElement!: Element;

  private rowsMap!: Map<number, Element>;
  private cellsMap!: Map<string, Element>;

  public fromXML(xmlString: string) {
    this.xmlDoc = new DOMParser().parseFromString(xmlString, "text/xml");

    this.worksheetElement = this.xmlDoc.getElementsByTagName('worksheet')[0];
    this.namespace = this.worksheetElement.getAttribute('xmlns') ?? "";

    this.sheetDataElement = this.xmlDoc.getElementsByTagName("sheetData")[0];
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

  public addRows(rows: Array<Array<CellObject>>, rowStartIndex: number, columnStartIndex: number): void {
    for(let i = 0; i < rows.length; i++) {
      const rowIndex = rowStartIndex + i;
      const row = rows[i];

      let rowElement = this.rowsMap.get(rowIndex);
      if(!rowElement) {
        rowElement = this.xmlDoc.createElementNS(this.namespace, 'row');
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
          cellElement = this.xmlDoc.createElementNS(this.namespace, 'c');
          cellElement.setAttribute('r', cellReference);
        }
        
        const cellStyle = row[j].Format.Style;
        if(cellStyle) cellElement.setAttribute('s', cellStyle ?? "");

        const cellType = row[j].Format.Type;
        if(cellType) cellElement.setAttribute('t', cellType ?? "");
        
        const valueElement = this.xmlDoc.createElementNS(this.namespace, 'v');
        valueElement.textContent = row[j].Value.toString();
        cellElement.replaceChildren(valueElement);
        rowElement.appendChild(cellElement);
      }
    }
  }

  public toString(): string {
    return new XMLSerializer().serializeToString(this.xmlDoc);
  }
}