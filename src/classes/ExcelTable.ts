import { CellIndex } from "../types";
import ExcelColumnConverter from "./ExcelColumnConverter";

export default class ExcelTable {
    private xmlDocument!: Document;
    private tableElement!: Element;
    public minCellIndex!: CellIndex;
    public maxCellIndex!: CellIndex;
    private autoFilterElement!: Element;
  
    public fromXML(xmlString: string) {
      this.xmlDocument = new DOMParser().parseFromString(xmlString, "text/xml");
      this.tableElement = this.xmlDocument.getElementsByTagName('table')[0];
      const [minCellRef, maxCellRef] = this.tableElement.getAttribute('ref')?.split(':') ?? ['A1', 'A1'];
      this.minCellIndex = ExcelColumnConverter.cellRefToIndex(minCellRef);
      this.maxCellIndex = ExcelColumnConverter.cellRefToIndex(maxCellRef);
      this.autoFilterElement = this.tableElement.getElementsByTagName('autoFilter')[0];
    }
  
    public toString(): string {
      const minCellRef = ExcelColumnConverter.numberToColumn(this.minCellIndex.columnIndex) + this.minCellIndex.rowIndex;
      const maxCellRef = ExcelColumnConverter.numberToColumn(this.maxCellIndex.columnIndex) + this.maxCellIndex.rowIndex;
      const tableRef = `${minCellRef}:${maxCellRef}`;
  
      this.tableElement.setAttribute('ref', tableRef);
      this.autoFilterElement.setAttribute('ref', tableRef); 
  
      return new XMLSerializer().serializeToString(this.xmlDocument);
    }
  }
  