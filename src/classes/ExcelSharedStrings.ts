export default class ExcelSharedStrings {
    private xmlDocument!: XMLDocument;
    private sstElement!: Element | null;
    private namespace!: string;
    public sharedStringArray!: Array<string | null | undefined>;

    public fromXML(xmlString: string) {
        this.xmlDocument = new DOMParser().parseFromString(xmlString, 'text/xml');
        this.sstElement = this.xmlDocument.querySelector('sst'); 
        if(!this.sstElement) throw new Error('Invalid sst element');
        this.namespace = this.sstElement.getAttribute('xmlns') ?? '';
        this.sharedStringArray = [];
        
        const siElements = this.sstElement.querySelectorAll('si');
        for(let i = 0; i < siElements.length; i++) {
            const tElement = siElements[i].querySelector('t');
            this.sharedStringArray[i] = tElement?.textContent ;
        }
    }

    public toString(): string {
        this.sstElement?.setAttribute('uniqueCount', this.sharedStringArray.length.toString()); 

        return new XMLSerializer().serializeToString(this.xmlDocument);
    }

    public getStringIndex(string: string): number {
        let stringIndex = this.sharedStringArray.findIndex((item) => item === string); 

        if(stringIndex === -1) {
            stringIndex = this.sharedStringArray.length; 
            this.sharedStringArray[stringIndex] = string; 

            const siElement = this.xmlDocument.createElementNS(this.namespace, 'si');
            const tElement = this.xmlDocument.createElementNS(this.namespace, 't');
            tElement.textContent = string;
            siElement.appendChild(tElement);
            this.sstElement?.appendChild(siElement);
        }
        return stringIndex
    }

    public getIndexString (index: number): string {
        return this.sharedStringArray[index] ?? '';
    }

    public getStringValues (): string {
        return JSON.stringify(this.sharedStringArray);
    }
}