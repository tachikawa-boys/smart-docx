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
import { Paragraph } from 'docx';
import { Header } from 'docx/build/file/header/header';

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

        doc.writeSection(paragraphs, null, null);
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

        doc.writeSection(paragraphs, null, null);
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

        doc.writeSection(paragraphs, null, null);
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

    it("Simple Header Write Verify", async () => {
        const temp = tmpdir;
        const doc = new DocWriter("HeaderTest", "TestDesc", "HeaderTestFile", "Test");
        doc.registerHandle(temp);

        const paragraph1 = new docx.Paragraph("HelloWorld!");
        const paragraphs: docx.Paragraph[] = new Array(paragraph1);
        const headerParagraph = new docx.Paragraph("HelloWorldHeader!");
        const header = new docx.Header();
        header.options.children.push(headerParagraph);

        doc.writeSection(paragraphs, header, null);
        await doc.closeHandle();

        const expectedFile = temp + '\\HeaderTestFile.docx';
        expect(file(expectedFile)).to.exist;
        const docR = new DocReader(expectedFile);

        docR.openDoc();
        expect(docR.searchForHeaderText("HelloWorldHeader!")).to.be.true;

        await docR.closeDoc();
        await cleanupTmp(expectedFile);
    });

    it("Simple Footer Write Verify", async () => {
        const temp = tmpdir;
        const doc = new DocWriter("FooterTest", "TestDesc", "FooterTestFile", "Test");
        doc.registerHandle(temp);

        const paragraph1 = new docx.Paragraph("HelloWorld!");
        const paragraphs: docx.Paragraph[] = new Array(paragraph1);
        const footerParagraph = new docx.Paragraph("HelloWorldFooter!");
        const footer = new docx.Footer();
        footer.options.children.push(footerParagraph);

        doc.writeSection(paragraphs, null, footer);
        await doc.closeHandle();

        const expectedFile = temp + '\\FooterTestFile.docx';
        expect(file(expectedFile)).to.exist;
        const docR = new DocReader(expectedFile);

        docR.openDoc();
        expect(docR.searchForFooterText("HelloWorldFooter!")).to.be.true;

        await docR.closeDoc();
        await cleanupTmp(expectedFile);
    });
});
