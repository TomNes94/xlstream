"use strict";
var __importDefault = (this && this.__importDefault) || function (mod) {
    return (mod && mod.__esModule) ? mod : { "default": mod };
};
Object.defineProperty(exports, "__esModule", { value: true });
var ssf_1 = __importDefault(require("ssf"));
var stream_1 = require("stream");
var StreamZip = require("node-stream-zip");
var saxStream = require("sax-stream");
function lettersToNumber(letters) {
    return letters.split("").reduce(function (r, a) { return r * 26 + parseInt(a, 36) - 9; }, 0);
}
function getXlsxStream(options) {
    return new Promise(function (resolve, reject) {
        function getTransform(formats, strings) {
            return new stream_1.Transform({
                objectMode: true,
                transform: function (chunk, encoding, done) {
                    var arr = [];
                    var formattedArr = [];
                    var obj = {};
                    var formattedObj = {};
                    var parsingHeader = false;
                    if (options.skipRows) {
                        if (options.skipRows.includes(parseInt(chunk.attribs.r))) {
                            done(undefined, null);
                            return;
                        }
                    }
                    var children = chunk.children ? (chunk.children.c.length ? chunk.children.c : [chunk.children.c]) : [];
                    for (var i = 0; i < children.length; i++) {
                        var ch = children[i];
                        if (ch.children) {
                            var value = ch.children.v.value;
                            if (ch.attribs.t === "s") {
                                value = strings[value];
                            }
                            //value = isNaN(value) ? value : Number(value);
                            var column = ch.attribs.r.replace(/[0-9]/g, "");
                            var index = lettersToNumber(column) - 1;
                            if (options.withHeader) {
                                if (!parsingHeader && header.length) {
                                    column = header[index];
                                }
                                else {
                                    header[index] = value;
                                    parsingHeader = true;
                                }
                            }
                            arr[index] = value;
                            obj[column] = value;
                            var formatId = ch.attribs.s ? Number(ch.attribs.s) : 0;
                            if (formatId) {
                                value = ssf_1.default.format(formats[formatId], value);
                                value = isNaN(value) ? value : Number(value);
                            }
                            formattedArr[index] = value;
                            formattedObj[column] = value;
                        }
                    }
                    done(undefined, parsingHeader || (options.ignoreEmpty && !arr.length)
                        ? null
                        : {
                            raw: {
                                obj: obj,
                                arr: arr,
                            },
                            formatted: {
                                obj: formattedObj,
                                arr: formattedArr,
                            },
                            header: header,
                        });
                },
            });
        }
        function processSheet(sheetId, formats, strings) {
            zip.stream("xl/worksheets/sheet" + sheetId + ".xml", function (err, stream) {
                var readStream = stream
                    .pipe(saxStream({
                    strict: true,
                    tag: "row",
                }))
                    .pipe(getTransform(formats, strings));
                stream.on("end", function () {
                    zip.close();
                });
                resolve(readStream);
            });
        }
        function processSharedStrings(sheetId, numberFormats, formats) {
            var strings = [];
            for (var i = 0; i < formats.length; i++) {
                var format = numberFormats[formats[i]];
                if (format) {
                    formats[i] = format;
                }
            }
            zip.stream("xl/sharedStrings.xml", function (err, stream) {
                if (stream) {
                    stream
                        .pipe(saxStream({
                        strict: true,
                        tag: "si",
                    }))
                        .on("data", function (x) {
                        if (x.children.t) {
                            strings.push(x.children.t.value);
                        }
                        else {
                            var str = "";
                            for (var i = 0; i < x.children.r.length; i++) {
                                var ch = x.children.r[i].children;
                                str += ch.t.value;
                            }
                            strings.push(str);
                        }
                    });
                    stream.on("end", function () {
                        processSheet(sheetId, formats, strings);
                    });
                }
                else {
                    processSheet(sheetId, formats, strings);
                }
            });
        }
        function processStyles(sheetId) {
            zip.stream("xl/styles.xml", function (err, stream) {
                var numberFormats = {};
                var formats = [];
                stream
                    .pipe(saxStream({
                    strict: true,
                    tag: ["cellXfs", "numFmts"],
                }))
                    .on("data", function (x) {
                    if (x.tag === "numFmts") {
                        var children = x.record.children.numFmt.length ? x.record.children.numFmt : [x.record.children.numFmt];
                        for (var i = 0; i < children.length; i++) {
                            numberFormats[Number(children[i].attribs.numFmtId)] = children[i].attribs.formatCode;
                        }
                    }
                    else if (x.tag === "cellXfs") {
                        for (var i = 0; i < x.record.children.xf.length; i++) {
                            var ch = x.record.children.xf[i];
                            formats[i] = Number(ch.attribs.numFmtId);
                        }
                    }
                });
                stream.on("end", function () {
                    processSharedStrings(sheetId, numberFormats, formats);
                });
            });
        }
        function processWorkbook() {
            zip.stream("xl/workbook.xml", function (err, stream) {
                var sheets = [];
                stream
                    .pipe(saxStream({
                    strict: true,
                    tag: "sheet",
                }))
                    .on("data", function (x) {
                    var attribs = x.attribs;
                    sheets.push(attribs.name);
                });
                stream.on("end", function () {
                    if (typeof options.sheet === "number") {
                        processStyles("" + (options.sheet + 1));
                    }
                    else if (typeof options.sheet === "string") {
                        processStyles("" + (sheets.indexOf(options.sheet) + 1));
                    }
                });
            });
        }
        var header = [];
        var zip = new StreamZip({
            file: options.filePath,
            storeEntries: true,
        });
        zip.on("ready", function () {
            processWorkbook();
        });
        zip.on("error", function (err) {
            reject(new Error(err));
        });
    });
}
exports.getXlsxStream = getXlsxStream;
function getWorksheets(options) {
    return new Promise(function (resolve, reject) {
        function processWorkbook() {
            zip.stream("xl/workbook.xml", function (err, stream) {
                stream
                    .pipe(saxStream({
                    strict: true,
                    tag: "sheet",
                }))
                    .on("data", function (x) {
                    sheets.push(x.attribs.name);
                });
                stream.on("end", function () {
                    zip.close();
                    resolve(sheets);
                });
            });
        }
        var sheets = [];
        var zip = new StreamZip({
            file: options.filePath,
            storeEntries: true,
        });
        zip.on("ready", function () {
            processWorkbook();
        });
    });
}
exports.getWorksheets = getWorksheets;
