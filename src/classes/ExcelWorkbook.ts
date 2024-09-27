import jsDOM from './SharedParserSerializer';

export default class ExcelTable {
    private xmlDocument!: Document;
    public sheets!: Map<number, Element>;
  
    public fromXML(xmlString: string) {
      this.xmlDocument = jsDOM.parser.parseFromString(xmlString, "text/xml");

      const sheets = this.xmlDocument.getElementsByTagName('sheet');
      for (let i = 1; i <= sheets.length; i++) {
        this.sheets.set(i, sheets[i]);
      }

    }
  
    public toString(): string {
      return jsDOM.serializer.serializeToString(this.xmlDocument);
    }
  }
  