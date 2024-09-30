export default class ExcelWorkbook {
    private xmlDocument!: Document;
    public sheets!: Array<string>;
  
    public fromXML(xmlString: string) {
      this.xmlDocument = new DOMParser().parseFromString(xmlString, "text/xml");

      const sheets = this.xmlDocument.getElementsByTagName('sheet');
      this.sheets = [];
      for (let i = 0; i < sheets.length; i++) {
        const sheet = sheets[i];
        const sheetName = sheet.getAttribute('name');
        if(!sheetName) throw new Error('A sheetName is null or invalid');
        this.sheets[i] = sheetName;
      }

    }
  
    public toString(): string {
      return new XMLSerializer().serializeToString(this.xmlDocument);
    }
  }
  