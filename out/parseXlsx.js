"use strict";
Object.defineProperty(exports, "__esModule", { value: true });
exports.ParseExcel = void 0;
var tsbeautify_1 = require("@brandless/tsbeautify");
var fs_1 = require("fs");
var node_xlsx_1 = require("node-xlsx");
var path_1 = require("path");
// const { default: dtsgenerator, parseSchema } = require('dtsgenerator');
/*
 * @Author: myx
 * @Date: 2022-05-16 16:40:06
 * @Last Modified by: myx
 * @Last Modified time: 2022-05-17 19:48:18
 */
var ParseExcel = /** @class */ (function () {
    function ParseExcel() {
        /** 解析文件格式 */
        this._strFormat = '.xlsx';
        /** 排除目录 */
        this._expired = [];
        /** 解析目录 */
        this._parseFile = './config_xlsx';
        /** json输出目录 */
        this._outJsonFile = './config_json';
        /** .d.ts输出目录 */
        this._outDtsFile = './config_dts';
        /** .d.ts内容 */
        this._dtsContext = '';
        /** 字段映射 */
        this._typeMap = {
            'INT': 'number',
            'STRING': 'string',
            'ARR': '[]',
            'ARRS': '[][]'
        };
    }
    ParseExcel.prototype.start = function () {
        this._dtsContext = "declare namespace config{\n";
        var files = this.getAllDoc(this._parseFile);
        for (var i = 0; i < files.length; i++) {
            var file = node_xlsx_1.default.parse(files[i]);
            this.parse(file);
        }
        this._dtsContext += "}";
        var tsbeautify = new tsbeautify_1.TsBeautifier();
        var result = tsbeautify.Beautify(this._dtsContext);
        (0, fs_1.writeFileSync)("".concat(this._outDtsFile, "/config.d.ts"), result);
    };
    /** 处理单张xlsx */
    ParseExcel.prototype.parse = function (file) {
        if (!file)
            return;
        var len = file.length;
        for (var i = 0; i < len; i++) {
            var child = file[i];
            //解析名字
            var name_1 = child.name;
            var arr = name_1.split('-');
            var prefix = arr[0];
            var suffix = arr[1];
            if (!prefix || !suffix)
                continue;
            console.log('解析-', name_1);
            var data = child.data;
            //第一行字段ID name value(根据本行内容长度生成解析长度)
            //第二行类型INT STRING STRING
            //第三行注释ID name 参数值 描述
            //第四行数据1 attack_shock {offset:10,time:0.1}
            var key = data[0];
            var type = data[1];
            var comment = data[2];
            var jsonContext = [];
            for (var m = 3; m < data.length; m++) {
                var p = data[m];
                if (p.length > 0) {
                    var foo = {};
                    for (var n = 0; n < key.length; n++) {
                        foo[key[n]] = p[n];
                    }
                    jsonContext.push(foo);
                }
            }
            //写入d.ts
            this._dtsContext += "/** ".concat(prefix, " */\n");
            this._dtsContext += "interface ".concat(suffix, "{\n");
            for (var m = 0; m < key.length; m++) {
                this._dtsContext += "/** ".concat(comment[m], " */\n");
                this._dtsContext += "".concat(key[m], ":").concat(this._typeMap[type[m]], ";\n");
            }
            this._dtsContext += "}\n";
            //写入json
            (0, fs_1.writeFileSync)("".concat(this._outJsonFile, "/").concat(suffix, ".json"), JSON.stringify(jsonContext));
        }
    };
    /** 获取目录下所有指定格式的文件 */
    ParseExcel.prototype.getAllDoc = function (mypath) {
        var _this = this;
        var result = [];
        var items = (0, fs_1.readdirSync)(mypath);
        items.map(function (item) {
            var temp = (0, path_1.join)(mypath, item);
            //文件夹则递归遍历
            if ((0, fs_1.statSync)(temp).isDirectory()) {
                _this.getAllDoc(temp);
                result = result.concat(_this.getAllDoc(temp));
            }
            else {
                var ext = (0, path_1.extname)(temp);
                if (ext.toLowerCase() === _this._strFormat) {
                    var base = (0, path_1.basename)(temp, _this._strFormat);
                    // if (this._expired.indexOf(base) >= 0)
                    result.push(temp);
                }
            }
        });
        return result;
    };
    return ParseExcel;
}());
exports.ParseExcel = ParseExcel;
var parseExcel = new ParseExcel();
parseExcel.start();
