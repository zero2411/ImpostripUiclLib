"use strict";
exports.__esModule = true;
exports.ImpostripUICL = exports.Unit = exports.BindingSide = exports.Sides = exports.CollatingMethod = void 0;
var convert = require("xml-js");
var CollatingMethod;
(function (CollatingMethod) {
    CollatingMethod["PerfectBound"] = "Perfect Bound";
    CollatingMethod["SaddleStiched"] = "Saddle Stiched";
    CollatingMethod["StepAndRepeat"] = "Step and Repeat";
    CollatingMethod["CutAndStack"] = "Cut and Stack";
    CollatingMethod["WebSectioning"] = "Web sectioning";
    CollatingMethod["RibbonCutAndStack"] = "Ribbon cut and stack";
    CollatingMethod["BalancedSections"] = "Balanced sections";
    CollatingMethod["ComingAndGoing"] = "Coming and Going";
})(CollatingMethod = exports.CollatingMethod || (exports.CollatingMethod = {}));
var Sides;
(function (Sides) {
    Sides["Simplex"] = "simplex";
    Sides["Duplex"] = "duplex";
})(Sides = exports.Sides || (exports.Sides = {}));
var BindingSide;
(function (BindingSide) {
    BindingSide["Left"] = "left";
    BindingSide["Right"] = "right";
    BindingSide["Top"] = "top";
    BindingSide["Bottom"] = "bottom";
})(BindingSide = exports.BindingSide || (exports.BindingSide = {}));
var Unit;
(function (Unit) {
    Unit["mm"] = "MM";
    Unit["cm"] = "CM";
    Unit["inch"] = "INCH";
    Unit["pt"] = "ADOBE";
})(Unit = exports.Unit || (exports.Unit = {}));
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
        };
        // Set Trimbox
        this.setTrimbox = function (width, height, unit) {
            this.trimbox = {
                width: width,
                height: height,
                unit: unit
            };
        };
        // Set Sheet
        this.setSheet = function (width, height, unit) {
            this.sheet = {
                width: width,
                height: height,
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
            this.margins = {
                top: top,
                bottom: bottom,
                left: left,
                right: right,
                unit: unit
            };
        };
        // Set Gutters
        this.setGutters = function (rowGutter, columnGutter, unit) {
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
                declarationKey: "xmlDeclaration"
            };
            var impo = {
                xmlDeclaration: {
                    attributes: {
                        version: "1.0",
                        encoding: "UTF-8"
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
                            Name: this.template
                        },
                        OtherSettings: {
                            attributes: {
                                DefaultPDFBox: "trimbox",
                                PDFPageAlignment: "center",
                                BuiltinMarkOnBlackOnly: "false",
                                NoPhaseInfoInOutputPSfile: "false",
                                GroupSignaturesOutput: "false",
                                GroupSignaturesAmount: "0",
                                PDFXObjectOptimization: "true"
                            }
                        }
                    },
                    PrintJob: {
                        Output: {
                            attributes: {
                                Format: "PDF",
                                Mockup: false,
                                PDFEngine: "adobelib",
                                OutputPath: this.OutputPath,
                                ImposedFiles: "Jobsperfile",
                                BackSideFirst: "false",
                                LastSigFirst: "false",
                                Rotation: "0"
                            }
                        },
                        attributes: {
                            JobID: this.jobName
                        }
                    }
                }
            };
            // Add MarkProfiles
            var markProfiles = [];
            for (var i = 0; i < this.markProfiles.length; i++) {
                markProfiles.push({
                    attributes: { Name: this.markProfiles[i] }
                });
            }
            impo.ImpostripOnDemand.CreateJob["MarkProfiles"] = { Profile: markProfiles };
            // Add Row- and Column-Gutters
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
            impo.ImpostripOnDemand.PrintJob["Gutters"] = gutters;
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
            impo.ImpostripOnDemand.PrintJob["Documents"] = { DocFile: documents };
            return convert.js2xml(impo, xmlOptions);
        };
    }
    return ImpostripUICL;
}());
exports.ImpostripUICL = ImpostripUICL;
