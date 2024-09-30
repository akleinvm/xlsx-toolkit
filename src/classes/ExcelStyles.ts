export default class ExcelStyles {
    private xmlDocument!: Document;
    private styleSheetElement!: Element;
    private namespace!: string;
    private cellFormatsElement!: Element;
    private cellFormatArray!: Array<number>;
  
    public fromXML(xmlString: string) {
      this.xmlDocument = new DOMParser().parseFromString(xmlString, "text/xml"); 
      this.styleSheetElement = this.xmlDocument.getElementsByTagName('styleSheet')[0]; 
      this.namespace = this.styleSheetElement.getAttribute('xmlns') ?? ""; 
  
      this.cellFormatsElement = this.styleSheetElement.getElementsByTagName('cellXfs')[0]; 
      const formats = this.cellFormatsElement.getElementsByTagName('xf'); 
      
      this.cellFormatArray = [];
      for (let i = 0; i < formats.length; i++) {
        const format = formats[i];
        const formatId = Number(format.getAttribute('numFmtId'));
        this.cellFormatArray.push(formatId);
      }
    }
  
    public getFormatIndex(formatId: number): string {
      let formatIndex = this.cellFormatArray.indexOf(formatId).toString();
  
      if(formatIndex === '-1') {
        formatIndex = this.cellFormatArray.length.toString();
        this.cellFormatArray.push(formatId);
        const xfElement = this.xmlDocument.createElementNS(this.namespace, 'xf');
        xfElement.setAttribute('numFmtId', formatId.toString());
        this.cellFormatsElement.appendChild(xfElement);
      }
  
      return formatIndex
    }
  
    public toString() {
      return new XMLSerializer().serializeToString(this.xmlDocument);
    }
}