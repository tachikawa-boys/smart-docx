import * as docx from "docx";
import { writeFileSync } from "fs";

/**
 * Document-writing tool. 
 */
export class DocWriter {
    private _doc: docx.Document;
    private _path: string;
    creator: string;
    description: string;
    title: string;
    headlinefont: string;
    classification: string;

    constructor(creator: string, description: string, title: string, classification: string) {
        this.creator = creator;
        this.description = description;
        this.title = title;
        this.headlinefont = 'Goudy Old Style';
        this.classification = classification;
    }

    /**
     * Registers the filehandle based on the title and directory.
     *  And will allow writing to start, enabling other functionality.
     */
    registerHandle(path: string): void {
        this._path = path;

        this._doc = new docx.Document({}, {}, []);
    }

    /**
    * Closes the active Document writer.
    */
    async closeHandle(): Promise<void> {
        await docx.Packer.toBuffer(this._doc).then((buffer) => {
            writeFileSync(this._path + "\\"+ this.title + ".docx", buffer);
        });
    }

    /** 
     * Applies a payload of paragraphs to a new section. 
     */
    writeSection(paragraphs: docx.Paragraph[]): void {
        paragraphs.push(new docx.Paragraph({children: [new docx.PageBreak()]}))

        this._doc.addSection({
            properties: {/*Can be added as Params.*/ }, children: paragraphs 
        });
    }
}
