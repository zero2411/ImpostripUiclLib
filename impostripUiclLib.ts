import * as convert from "xml-js";

export enum CollatingMethod {
  PerfectBound = "Perfect Bound",
  SaddleStiched = "Saddle Stiched",
  StepAndRepeat = "Step and Repeat",
  CutAndStack = "Cut and Stack",
  WebSectioning = "Web sectioning",
  RibbonCutAndStack = "Ribbon cut and stack",
  BalancedSections = "Balanced sections",
  ComingAndGoing = "Coming and Going",
}

export enum Sides {
  Simplex = "simplex",
  Duplex = "duplex",
}

export enum BindingSide {
  Left = "left",
  Right = "right",
  Top = "top",
  Bottom = "bottom",
}

export enum Unit {
  mm = "MM",
  cm = "CM",
  inch = "INCH",
  pt = "ADOBE",
}
export class ImpostripUICL {
  private jobName: string;
  private outputPath: string;
  private trimbox: {};
  private sheet: {};
  private margins: {};
  private gutters: {};
  private rows: number;
  private columns: number;
  private bindingCollatingMethod: object;
  private bindingSides: string;
  private bindingSide: string;
  private markProfiles: string[] = [];
  private documents: {}[] = [];
  private errorMessages: string[] = [];

  // Set JobName
  setJobName = function (jobName: string) {
    this.jobName = jobName;
  };

  // Set OutputPath
  setOutputPath = function (outputPath: string) {
    this.outputPath = outputPath;
  };

  // Set CollatingMethod
  setBinding = function (bindingCollatingMethod: object, bindingSides: string, bindingSide: string) {
    this.bindingCollatingMethod = bindingCollatingMethod;
    this.bindingSides = bindingSides;
    this.bindingSide = bindingSide;
  };

  // Set Trimbox
  setTrimbox = function (width: number, height: number, unit: Unit) {
    this.trimbox = {
      width: width,
      height: height,
      unit: unit,
    };
  };

  // Set Sheet
  setSheet = function (width: number, height: number, unit: Unit) {
    this.sheet = {
      width: width,
      height: height,
      unit: unit,
    };
  };

  // Set Rows and Columns
  setRowsAndColumns = function (rows: number, columns: number) {
    this.rows = rows;
    this.columns = columns;
  };

  // Set Margins
  setMargins = function (top: number, bottom: number, left: number, right: number, unit: Unit) {
    this.margins = {
      top: top,
      bottom: bottom,
      left: left,
      right: right,
      unit: unit,
    };
  };

  // Set Gutters
  setGutters = function (rowGutter: number, columnGutter: number, unit: Unit) {
    this.gutters = {
      rowGutter: rowGutter,
      columnGutter: columnGutter,
      unit: unit,
    };
  };

  // Set Template -
  /**
   * Can be used to overide the calculated template
   *
   * @param template {string}
   */
  setTemplate = function (template: string) {
    this.template = template;
  };

  // Add MarkPofile
  /**
   * Add a markprofile
   *
   * @param markprofile {string}
   */
  addMarkProfile = function (markProfile: string) {
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
  addDocument = function (fullPathName: string, useAllPages: boolean = true, start: number = 0, end: number = 0) {
    this.documents.push({
      fullPathName: fullPathName,
      useAllPages: useAllPages,
      start: start,
      end: end,
    });
  };
  checkSettings = function (): boolean {
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
  getXML = function () {
    // Checks

    if (!this.checkSettings()) {
      throw new Error(this.errorMessages.join(" | "));
    }

    // Set Template
    if (this.template == undefined) {
      this.template = this.bindingCollatingMethod.code + "_" + this.rows + "x" + this.columns + "_" + this.bindingSides;
    }

    let xmlOptions = {
      compact: true,
      attributesKey: "attributes",
      ignoreComment: false,
      spaces: 4,
      declarationKey: "xmlDeclaration",
    };

    let impo = {
      xmlDeclaration: {
        attributes: {
          version: "1.0",
          encoding: "UTF-8",
        },
      },
      ImpostripOnDemand: {
        CreateJob: {
          attributes: {
            RefJobName: this.jobName,
          },
          Binding: {
            attributes: {
              CollatingMethod: this.bindingCollatingMethod.name,
              Sides: this.bindingSides,
              BindingSide: this.bindingSide,
            },
          },
          TrimSize: {
            attributes: {
              Dynamic: "false",
              Width: this.trimbox.width,
              Height: this.trimbox.height,
              Unit: this.trimbox.unit,
            },
          },
          PaperSize: {
            attributes: {
              Width: this.sheet.width,
              Height: this.sheet.height,
              Unit: this.sheet.unit,
            },
          },
          Template: {
            Name: this.template,
          },
          OtherSettings: {
            attributes: {
              DefaultPDFBox: "trimbox",
              PDFPageAlignment: "center",
              BuiltinMarkOnBlackOnly: "false",
              NoPhaseInfoInOutputPSfile: "false",
              GroupSignaturesOutput: "false",
              GroupSignaturesAmount: "0",
              PDFXObjectOptimization: "true",
            },
          },
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
              Rotation: "0",
            },
          },
          attributes: {
            JobID: this.jobName,
          },
        },
      },
    };

    // Add MarkProfiles
    let markProfiles = [];
    for (let i = 0; i < this.markProfiles.length; i++) {
      markProfiles.push({
        attributes: { Name: this.markProfiles[i] },
      });
    }
    impo.ImpostripOnDemand.CreateJob["MarkProfiles"] = { Profile: markProfiles };

    // Add Row- and Column-Gutters
    let rowGutters = [];
    for (let i = 1; i < this.rows; i++) {
      rowGutters.push({
        attributes: { Index: i, Value: this.gutters.rowGutter, Unit: this.gutters.unit },
      });
    }
    let columnGutters = [];
    for (let i = 1; i < this.columns; i++) {
      columnGutters.push({
        attributes: { Index: i, Value: this.gutters.columnGutter, Unit: this.gutters.unit },
      });
    }
    let gutters = {
      RowGutter: rowGutters,
      ColumnGutter: columnGutters,
    };
    impo.ImpostripOnDemand.PrintJob["Gutters"] = gutters;

    // Add documents
    let documents = [];
    for (let i = 0; i < this.documents.length; i++) {
      let document = {
        FullPathName: this.documents[i].fullPathName,
        PageRange: { attributes: { UseAllPages: this.documents[i].useAllPages } },
      };
      if (!this.documents[i].useAllPages) {
        document.PageRange.attributes["Start"] = this.documents[i].start;
        document.PageRange.attributes["End"] = this.documents[i].end;
      }
      documents.push(document);
    }
    impo.ImpostripOnDemand.PrintJob["Documents"] = { DocFile: documents };

    return convert.js2xml(impo, xmlOptions);
  };
}
