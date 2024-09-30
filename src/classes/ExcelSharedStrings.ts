export default class ExcelSharedStrings {
    private xmlDocument!: XMLDocument;
    private sstElement!: Element;
    private namespace!: string;
    public stringsMap!: Map<string, number>;
    public numbersMap!: Map<number, string>;

    public fromXML(xmlString: string) {
        this.xmlDocument = new DOMParser().parseFromString(xmlString, 'text/xml');
        this.sstElement = this.xmlDocument.getElementsByTagName('sst')[0];
        this.namespace = this.sstElement.getAttribute('xmlns') ?? '';
        this.stringsMap = new Map();
        this.numbersMap = new Map();
        
        const siElements = this.sstElement.getElementsByTagName('si');
        for(let i = 0; i < siElements.length; i++) {
            const tElement = siElements[i].getElementsByTagName('t')[0];
            if(tElement && tElement.textContent) {
                this.stringsMap.set(tElement.textContent, i);
                this.numbersMap.set(i, tElement.textContent);
            }
        }
    }

    public toString(): string {
        this.sstElement.setAttribute('uniqueCount', this.stringsMap.size.toString());

        return new XMLSerializer().serializeToString(this.xmlDocument);
    }

    public getStringIndex(string: string): number {
        let stringIndex = this.stringsMap.get(string);

        if(!stringIndex) {
            stringIndex = this.stringsMap.size;
            this.stringsMap.set(string, this.stringsMap.size);
            this.numbersMap.set(this.stringsMap.size, string);
            const siElement = this.xmlDocument.createElementNS(this.namespace, 'si',);
            const tElement = this.xmlDocument.createElementNS(this.namespace, 't');
            tElement.textContent = string;
            siElement.appendChild(tElement);
            this.sstElement.appendChild(siElement);
        }
        return stringIndex
    }

    public getIndexString (index: number): string {
        return this.numbersMap.get(index) ?? '';
    }
}