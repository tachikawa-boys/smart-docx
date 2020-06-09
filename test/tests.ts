import chai = require('chai');
import chaiFS = require('chai-fs');
import chaiFiles = require('chai-files');
import * as docx from "docx";
import { file } from "chai-files";
import { unlink } from "fs";
import { tmpdir } from 'tmp';
import { DocReader } from "../ts-src/DocReader";
import { DocWriter } from "../ts-src/DocWriter";
import randomstring = require('randomstring');

const expect = chai.expect;
const file = chaiFiles.file;
chai.use(chaiFS);
chai.use(chaiFiles);

/**
 * Helper function for cleaning up the file.
 * @param expectedFile
 */
async function cleanupTmp(expectedFile: string) {
    await unlink(expectedFile, (err) => {
        if (err) {
            throw err;
        }
        expect(file(expectedFile)).to.not.exist;
    });
}



describe("Basic Unit Tests", () => {
    it("Create docx", async () => {

        const randFileName = randomstring.generate(7);
        const randString = randomstring.generate();

        const temp = tmpdir;
        const doc = new DocWriter("Test", "TestDesc", randFileName, "Test");
        doc.registerHandle(temp);
        await doc.closeHandle();

        //const expectedFile = temp + '\\TestFile.docx';
        const expectedFile = temp + '\\' + randFileName + '.docx';
        expect(file(expectedFile)).to.exist;
        await cleanupTmp(expectedFile);
    });

    it("Simple Write no Verify", async () => {
        const randFileName = randomstring.generate(7);
        const randString = randomstring.generate();
        const temp = tmpdir;
        const doc = new DocWriter("Test", "TestDesc", randFileName, "Test");
        doc.registerHandle(temp);

        const paragraph1 = new docx.Paragraph(randString);
        const paragraphs: docx.Paragraph[] = new Array(paragraph1);

        doc.writeSection(paragraphs);
        await doc.closeHandle();

        //const expectedFile = temp + '\\TestFile.docx';
        const expectedFile = temp + '\\' + randFileName + '.docx';
        await cleanupTmp(expectedFile);
    });
});

describe("Read/Write Unit Tests", () => {

    it("Simple Write Verify", async () => {
        const randFileName = randomstring.generate(7);
        const randString = randomstring.generate();
        const temp = tmpdir;
        const doc = new DocWriter("Test", "TestDesc", randFileName, "Test");
        doc.registerHandle(temp);

        const paragraph1 = new docx.Paragraph(randString);
        const paragraphs: docx.Paragraph[] = new Array(paragraph1);

        doc.writeSection(paragraphs);
        await doc.closeHandle();

        //const expectedFile = temp + '\\TestFile.docx';
        const expectedFile = temp + '\\' + randFileName + '.docx';
        expect(file(expectedFile)).to.exist;
        const docR = new DocReader(expectedFile);

        docR.openDoc();
        expect(docR.searchForBodyText(randString)).to.be.true;
        await docR.closeDoc();
        await cleanupTmp(expectedFile);
    });

    it("Simple Write Verify (Fail)", async () => {
        const randFileName = randomstring.generate(7);
        const randString = randomstring.generate();
        const randStringFail = randomstring.generate(31);
        const temp = tmpdir;
        const doc = new DocWriter("Test", "TestDesc", randFileName, "Test");
        doc.registerHandle(temp);

        const paragraph1 = new docx.Paragraph(randString);
        const paragraphs: docx.Paragraph[] = new Array(paragraph1);

        doc.writeSection(paragraphs);
        await doc.closeHandle();


        //const expectedFile = temp + '\\TestFile.docx';
        const expectedFile = temp + '\\' + randFileName + '.docx';
        expect(file(expectedFile)).to.exist;
        const docR = new DocReader(expectedFile);

        docR.openDoc();
        expect(docR.searchForBodyText(randStringFail)).to.be.false;
        await docR.closeDoc();
        await cleanupTmp(expectedFile);
    });
});
