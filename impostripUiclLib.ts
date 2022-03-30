import * as convert from "xml-js";

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

export const CollatingMethod = {
  PerfectBound: { name: "Perfect Bound", code: "PB" },
  SaddleStiched: { name: "Saddle Stiched", code: "SS" },
  StepAndRepeat: { name: "Step and Repeat", code: "SR" },
  CutAndStack: { name: "Cut and Stack", code: "CS" },
  WebSectioning: { name: "Web sectioning", code: "WS" },
  RibbonCutAndStack: { name: "Ribbon cut and stack", code: "RCS" },
  BalancedSections: { name: "Balanced sections", code: "BS" },
  ComingAndGoing: { name: "Coming and Going", code: "CG" },
} as const;

export const Sides = {
  Simplex: "simplex",
  Duplex: "duplex",
} as const;

export const BindingSide = {
  Left: "left",
  Right: "right",
  Top: "top",
  Bottom: "bottom",
} as const;

export const Unit = {
  mm: "MM",
  cm: "CM",
  inch: "INCH",
  pt: "ADOBE",
} as const;

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
    console.info(this);
  };

  // Set Trimbox
  setTrimbox = function (width: number, height: number, unit = Unit.mm) {
    this.trimbox = {
      width: Math.round(width * 10) / 10,
      height: Math.round(height * 10) / 10,
      unit: "MM",
    };
  };

  // Set Sheet
  setSheet = function (width: number, height: number, unit = Unit.mm) {
    this.sheet = {
      width: Math.round(width),
      height: Math.round(height),
      unit: unit,
    };
  };

  // Set Rows and Columns
  setRowsAndColumns = function (rows: number, columns: number) {
    this.rows = rows;
    this.columns = columns;
  };

  // Set Margins
  setMargins = function (top: number, bottom: number, left: number, right: number, unit = "MM") {
    this.margins = {
      top: Math.round(top * 10) / 10,
      bottom: Math.round(bottom * 10) / 10,
      left: Math.round(left * 10) / 10,
      right: Math.round(right * 10) / 10,
      unit: unit,
    };
  };

  // Set Gutters
  setGutters = function (rowGutter: number, columnGutter: number, unit = "MM") {
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
      declarationKey: 'xmlDeclaration',
    };

    let impo = {
      xmlDeclaration: {
        attributes: {
          version: "1.0",
          encoding: "UTF-8",
          standalone: "no",
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
            TrimSize: {
              attributes: {
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
          },
        },
        Marks: {
          attributes: {
            ShowFoldingMark: "false",
            ShowCutMark: "true",
          },
          CutMark: {
            DoubleCutMark: "false",
            QuarterInchCutMark: "false",
            TrimBoxCutMark: "true",
          },
        },
        Output: {
          attributes: {
            Format: "PDF",
            Mockup: false,
            PDFEngine: "adobelib",
            OutputPath: this.OutputPath,
            ImposedFiles: "Jobsperfile",
          },
        },
        OtherSettings: {
          attributes: {
            DefaultPDFBox: "Trimbox",
          },
        },
        PrintJob: {
          attributes: {
            JobID: this.jobName,
          },
          RefJob: {
            Name: this.jobName,
          },
        },
      },
    };

    // Add row- and column-gutters
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
    impo.ImpostripOnDemand.CreateJob.Template["Gutters"] = gutters;

    // Add marksets
    let markProfiles = [];
    for (let i = 0; i < this.markProfiles.length; i++) {
      markProfiles.push({
        attributes: { Name: this.markProfiles[i] },
      });
    }
    impo.ImpostripOnDemand["MarkProfiles"] = { Profile: markProfiles };

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
    impo.ImpostripOnDemand["Documents"] = { DocFile: documents };

    return convert.js2xml(impo, xmlOptions);
  };
}

//test();
