
import { TsBeautifier } from '@brandless/tsbeautify';
import { readdirSync, statSync, writeFileSync } from 'fs';
import xlsx from 'node-xlsx';
import { basename, extname, join } from "path"
// const { default: dtsgenerator, parseSchema } = require('dtsgenerator');

/*
 * @Author: myx 
 * @Date: 2022-05-16 16:40:06 
 * @Last Modified by: myx
 * @Last Modified time: 2022-05-17 18:46:31
 */
export class ParseExcel {
    /** 解析文件格式 */
    private readonly _strFormat: string = '.xlsx';
    /** 排除目录 */
    private readonly _expired: string[] = [];
    /** 解析目录 */
    private readonly _parseFile: string = './config_xlsx';
    /** json输出目录 */
    private readonly _outJsonFile: string = './config_json';
    /** .d.ts输出目录 */
    private readonly _outDtsFile: string = './config_dts';
    /** .d.ts内容 */
    private _dtsContext: string = '';
    /** 字段映射 */
    private readonly _typeMap: { [key: string]: string } = {
        'INT': 'number',
        'STRING': 'string',
        'ARR': '[]',
        'ARRS': '[][]'
    };
    constructor() {
    }

    start(): void {
        this._dtsContext = `declare namespace xh.conf{\n`;
        const files = this.getAllDoc(this._parseFile);
        for (let i = 0; i < files.length; i++) {
            const file = xlsx.parse(files[i]);
            this.parse(file);
        }
        this._dtsContext += `}`
        let tsbeautify = new TsBeautifier();
        let result = tsbeautify.Beautify(this._dtsContext);
        writeFileSync(`${this._outDtsFile}/config.d.ts`, result);
    }

    /** 处理单张xlsx */
    private parse(file): void {
        if (!file) return;
        let len = file.length;
        for (let i = 0; i < len; i++) {
            const child = file[i];
            //解析名字
            const name: string = child.name;
            const arr = name.split('-');
            let prefix = arr[0];
            let suffix = arr[1];
            if (!prefix || !suffix) continue;
            console.log('解析-', name);
            const data = child.data;
            //第一行字段ID name value(根据本行内容长度生成解析长度)
            //第二行类型INT STRING STRING
            //第三行注释ID name 参数值 描述
            //第四行数据1 attack_shock {offset:10,time:0.1}
            let key = data[0];
            let type = data[1];
            let comment = data[2];
            let jsonContext = [];
            for (let m = 3; m < data.length; m++) {
                const p = data[m];
                if (p.length > 0) {
                    let foo = {};
                    for (let n = 0; n < key.length; n++) {
                        foo[key[n]] = p[n];
                    }
                    jsonContext.push(foo);
                }
            }
            //写入d.ts
            this._dtsContext += `/** ${prefix} */\n`;
            this._dtsContext += `interface ${suffix}{\n`;
            for (let m = 0; m < key.length; m++) {
                this._dtsContext += `/** ${comment[m]} */\n`;
                this._dtsContext += `${key[m]}:${this._typeMap[type[m]]};\n`;
            }
            this._dtsContext += `}\n`

            //写入json
            writeFileSync(`${this._outJsonFile}/${suffix}.json`, JSON.stringify(jsonContext))
        }
    }

    /** 获取目录下所有指定格式的文件 */
    private getAllDoc(mypath: string): string[] {
        let result: string[] = [];
        let items = readdirSync(mypath);
        items.map(item => {
            let temp = join(mypath, item);
            //文件夹则递归遍历
            if (statSync(temp).isDirectory()) {
                this.getAllDoc(temp)
                result = result.concat(this.getAllDoc(temp));
            } else {
                let ext = extname(temp);
                if (ext.toLowerCase() === this._strFormat) {
                    const base = basename(temp, this._strFormat);
                    // if (this._expired.indexOf(base) >= 0)
                    result.push(temp);
                }
            }
        })
        return result;
    }
}
const parseExcel = new ParseExcel();
parseExcel.start();

