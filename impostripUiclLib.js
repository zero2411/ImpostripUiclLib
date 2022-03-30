"use strict";
exports.__esModule = true;
exports.ImpostripUICL = exports.Unit = exports.BindingSide = exports.Sides = exports.CollatingMethod = void 0;
var convert = require("xml-js");
// function test() {
//   try {
//     let imp = new ImpostripUICL();
//     imp.setJobName("TestJob");
//     imp.setOutputPath("C:\\Users\\admin-wihabo\\Desktop\\Output");
//     imp.setBinding(CollatingMethod.CutAndStack , Sides.Duplex, BindingSide.Left);
//     imp.setTrimbox(171,141, "MM");
//     imp.setSheet(748, 528, "MM");
//     imp.setRowsAndColumns(4,4);
//     imp.setMargins(28.5,28.5,24.5,24.5, "MM");
//     imp.setGutters(5, 5, "MM");
//     imp.addMarkProfile("SF Cutmarks");
//     imp.addMarkProfile("SF Slugline");
//     imp.addDocument("C:\\Users\\admin-wihabo\\Desktop\\Test bestanden\\TEST 114x171 duplex\\Dummy 114x171 - Duplex - Cyaan.pdf");
//     imp.addDocument("C:\\Users\\admin-wihabo\\Desktop\\Test bestanden\\TEST 114x171 duplex\\Dummy 114x171 - Duplex - Geel.pdf");
//     imp.addDocument("C:\\Users\\admin-wihabo\\Desktop\\Test bestanden\\TEST 114x171 duplex\\Dummy 171x114 - Duplex - Groen.pdf");
//     imp.addDocument("C:\\Users\\admin-wihabo\\Desktop\\Test bestanden\\TEST 114x171 duplex\\Dummy 171x114 - Duplex - Paars.pdf");
//     console.log(imp.getXML());
//   } catch (e) {
//     console.error(e);
//   }
// }
exports.CollatingMethod = {
    PerfectBound: { name: "Perfect Bound", code: "PB" },
    SaddleStiched: { name: "Saddle Stiched", code: "SS" },
    StepAndRepeat: { name: "Step and Repeat", code: "SR" },
    CutAndStack: { name: "Cut and Stack", code: "CS" },
    WebSectioning: { name: "Web sectioning", code: "WS" },
    RibbonCutAndStack: { name: "Ribbon cut and stack", code: "RCS" },
    BalancedSections: { name: "Balanced sections", code: "BS" },
    ComingAndGoing: { name: "Coming and Going", code: "CG" }
};
exports.Sides = {
    Simplex: "simplex",
    Duplex: "duplex"
};
exports.BindingSide = {
    Left: "left",
    Right: "right",
    Top: "top",
    Bottom: "bottom"
};
exports.Unit = {
    mm: "MM",
    cm: "CM",
    inch: "INCH",
    pt: "ADOBE"
};
var ImpostripUICL = /** @class */ (function () {
    function ImpostripUICL() {
        this.markProfiles = [];
        this.documents = [];
        this.errorMessages = [];
        // Set JobName
        this.setJobName = function (jobName) {
            this.jobName = jobName;
        };
        // Set OutputPath
        this.setOutputPath = function (outputPath) {
            this.outputPath = outputPath;
        };
        // Set CollatingMethod
        this.setBinding = function (bindingCollatingMethod, bindingSides, bindingSide) {
            this.bindingCollatingMethod = bindingCollatingMethod;
            this.bindingSides = bindingSides;
            this.bindingSide = bindingSide;
            console.info(this);
        };
        // Set Trimbox
        this.setTrimbox = function (width, height, unit) {
            if (unit === void 0) { unit = exports.Unit.mm; }
            this.trimbox = {
                width: Math.round(width * 10) / 10,
                height: Math.round(height * 10) / 10,
                unit: "MM"
            };
        };
        // Set Sheet
        this.setSheet = function (width, height, unit) {
            if (unit === void 0) { unit = exports.Unit.mm; }
            this.sheet = {
                width: Math.round(width),
                height: Math.round(height),
                unit: unit
            };
        };
        // Set Rows and Columns
        this.setRowsAndColumns = function (rows, columns) {
            this.rows = rows;
            this.columns = columns;
        };
        // Set Margins
        this.setMargins = function (top, bottom, left, right, unit) {
            if (unit === void 0) { unit = "MM"; }
            this.margins = {
                top: Math.round(top * 10) / 10,
                bottom: Math.round(bottom * 10) / 10,
                left: Math.round(left * 10) / 10,
                right: Math.round(right * 10) / 10,
                unit: unit
            };
        };
        // Set Gutters
        this.setGutters = function (rowGutter, columnGutter, unit) {
            if (unit === void 0) { unit = "MM"; }
            this.gutters = {
                rowGutter: rowGutter,
                columnGutter: columnGutter,
                unit: unit
            };
        };
        // Set Template -
        /**
         * Can be used to overide the calculated template
         *
         * @param template {string}
         */
        this.setTemplate = function (template) {
            this.template = template;
        };
        // Add MarkPofile
        /**
         * Add a markprofile
         *
         * @param markprofile {string}
         */
        this.addMarkProfile = function (markProfile) {
            this.markProfiles.push(markProfile);
        };
        // Add Document
        /**
         * Add a document
         *
         * @param fullPathName {string}
         * @param [useAllPages] {boolean}
         * @param [start] {number}
         * @param [end] {number}
         */
        this.addDocument = function (fullPathName, useAllPages, start, end) {
            if (useAllPages === void 0) { useAllPages = true; }
            if (start === void 0) { start = 0; }
            if (end === void 0) { end = 0; }
            this.documents.push({
                fullPathName: fullPathName,
                useAllPages: useAllPages,
                start: start,
                end: end
            });
        };
        this.checkSettings = function () {
            if (this.jobName == undefined) {
                this.errorMessages.push("Jobname not set.");
            }
            if (this.outputPath == undefined) {
                this.errorMessages.push("OutputPath not set.");
            }
            if (this.trimbox == undefined) {
                this.errorMessages.push("Trimbox not set.");
            }
            if (this.sheet == undefined) {
                this.errorMessages.push("Sheet not set.");
            }
            if (this.rows == undefined) {
                this.errorMessages.push("Rows not set.");
            }
            if (this.columns == undefined) {
                this.errorMessages.push("Columns not set.");
            }
            if (this.documents.length == 0) {
                this.errorMessages.push("No documents added.");
            }
            if (this.errorMessages.length > 0) {
                console.log(this.errorMessages);
                return false;
            }
            return true;
        };
        // Get XML
        this.getXML = function () {
            // Checks
            if (!this.checkSettings()) {
                throw new Error(this.errorMessages.join(" | "));
            }
            // Set Template
            if (this.template == undefined) {
                this.template = this.bindingCollatingMethod.code + "_" + this.rows + "x" + this.columns + "_" + this.bindingSides;
            }
            var xmlOptions = {
                compact: true,
                attributesKey: "attributes",
                ignoreComment: false,
                spaces: 4,
                declarationKey: 'xmlDeclaration'
            };
            var impo = {
                xmlDeclaration: {
                    attributes: {
                        version: "1.0",
                        encoding: "UTF-8",
                        standalone: "no"
                    }
                },
                ImpostripOnDemand: {
                    CreateJob: {
                        attributes: {
                            RefJobName: this.jobName
                        },
                        Binding: {
                            attributes: {
                                CollatingMethod: this.bindingCollatingMethod.name,
                                Sides: this.bindingSides,
                                BindingSide: this.bindingSide
                            }
                        },
                        TrimSize: {
                            attributes: {
                                Dynamic: "false",
                                Width: this.trimbox.width,
                                Height: this.trimbox.height,
                                Unit: this.trimbox.unit
                            }
                        },
                        PaperSize: {
                            attributes: {
                                Width: this.sheet.width,
                                Height: this.sheet.height,
                                Unit: this.sheet.unit
                            }
                        },
                        Template: {
                            Name: this.template,
                            TrimSize: {
                                attributes: {
                                    Width: this.trimbox.width,
                                    Height: this.trimbox.height,
                                    Unit: this.trimbox.unit
                                }
                            },
                            PaperSize: {
                                attributes: {
                                    Width: this.sheet.width,
                                    Height: this.sheet.height,
                                    Unit: this.sheet.unit
                                }
                            }
                        }
                    },
                    Marks: {
                        attributes: {
                            ShowFoldingMark: "false",
                            ShowCutMark: "true"
                        },
                        CutMark: {
                            DoubleCutMark: "false",
                            QuarterInchCutMark: "false",
                            TrimBoxCutMark: "true"
                        }
                    },
                    Output: {
                        attributes: {
                            Format: "PDF",
                            Mockup: false,
                            PDFEngine: "adobelib",
                            OutputPath: this.OutputPath,
                            ImposedFiles: "Jobsperfile"
                        }
                    },
                    OtherSettings: {
                        attributes: {
                            DefaultPDFBox: "Trimbox"
                        }
                    },
                    PrintJob: {
                        attributes: {
                            JobID: this.jobName
                        },
                        RefJob: {
                            Name: this.jobName
                        }
                    }
                }
            };
            // Add row- and column-gutters
            var rowGutters = [];
            for (var i = 1; i < this.rows; i++) {
                rowGutters.push({
                    attributes: { Index: i, Value: this.gutters.rowGutter, Unit: this.gutters.unit }
                });
            }
            var columnGutters = [];
            for (var i = 1; i < this.columns; i++) {
                columnGutters.push({
                    attributes: { Index: i, Value: this.gutters.columnGutter, Unit: this.gutters.unit }
                });
            }
            var gutters = {
                RowGutter: rowGutters,
                ColumnGutter: columnGutters
            };
            impo.ImpostripOnDemand.CreateJob.Template["Gutters"] = gutters;
            // Add marksets
            var markProfiles = [];
            for (var i = 0; i < this.markProfiles.length; i++) {
                markProfiles.push({
                    attributes: { Name: this.markProfiles[i] }
                });
            }
            impo.ImpostripOnDemand["MarkProfiles"] = { Profile: markProfiles };
            // Add documents
            var documents = [];
            for (var i = 0; i < this.documents.length; i++) {
                var document_1 = {
                    FullPathName: this.documents[i].fullPathName,
                    PageRange: { attributes: { UseAllPages: this.documents[i].useAllPages } }
                };
                if (!this.documents[i].useAllPages) {
                    document_1.PageRange.attributes["Start"] = this.documents[i].start;
                    document_1.PageRange.attributes["End"] = this.documents[i].end;
                }
                documents.push(document_1);
            }
            impo.ImpostripOnDemand["Documents"] = { DocFile: documents };
            return convert.js2xml(impo, xmlOptions);
        };
    }
    return ImpostripUICL;
}());
exports.ImpostripUICL = ImpostripUICL;
//test();
