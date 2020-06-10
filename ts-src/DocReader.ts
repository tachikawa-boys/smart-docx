import fs = require('fs');
import unzip = require('adm-zip');
import tmp = require('tmp');
import xml2js = require('xml2js');
import glob = require('glob');

/**
 * Reader to break open an Open XML Word file to verify values.
 * */
export class DocReader {
    private _path: string;
    private _tmp: string;

    constructor(path: string) {
        this._path = path;
    }

    /**
     * Spins up a temporary directory and explodes the content of the Open XML into it.
     * */
    public openDoc() {

        const tmpobj = tmp.dirSync();
        this._tmp = tmpobj.name;
        console.log('Dir: ', tmpobj.name);
        tmpobj.removeCallback();

        const zipHandle = new unzip(this._path);
        zipHandle.extractAllTo(this._tmp, true);
    }

    /**
     * Simplest possible v
     * @param text - Any string contained within the XML file.
     * Since this is super-simple, it may misfire on XML namespacing.
     */
    public searchForBodyText(text: string): boolean {
        const data = fs.readFileSync(this._tmp + '\\word\\document.xml', { encoding: 'utf-8' });
        xml2js.parseString(data, function (err, res) {
            if (err) console.log(err);
            console.log(res);
        });
        if (data.includes(text)) {
            return true;
        };
        return false;
    }

    /**
     * Simplest possible method for searching text in headers.
     * @param text - Any string contained within the XML file.
     * Since this is super-simple, it may misfire on XML namespacing.
     */
    public searchForHeaderText(text: string): boolean {
        const files = glob.sync(this._tmp + '/word/header?.xml');
        const retVal = files.some((file) => {
            const data = fs.readFileSync(file, { encoding: 'utf-8' });
            xml2js.parseString(data, function (err, res) {
                if (err) console.log(err);
                console.log(res);
            });
            if (data.includes(text)) {
                return true;
            };
            return false;
        });
        return retVal;
    }

    /**
 * Simplest possible method for searching text in headers.
 * @param text - Any string contained within the XML file.
 * Since this is super-simple, it may misfire on XML namespacing.
 */
    public searchForFooterText(text: string): boolean {
        const files = glob.sync(this._tmp + '/word/footer?.xml');
        const retVal = files.some((file) => {
            const data = fs.readFileSync(file, { encoding: 'utf-8' });
            xml2js.parseString(data, function (err, res) {
                if (err) console.log(err);
                console.log(res);
            });
            if (data.includes(text)) {
                return true;
            };
            return false;
        });
        return retVal;
    }

    /**
     * Closes the open handle - and deletes the temporary directory.
     * Don't call this until you're completely done searching!
     * */
    public async closeDoc() {
        await fs.rmdir(this._tmp, (err) => {
            if (err) {
                throw err;
            }

            console.log(`${this._tmp} is deleted!`);
        });
    }

}



