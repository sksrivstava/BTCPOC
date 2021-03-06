﻿import * as React from "react";
import { connect } from 'react-redux';
import { RouteComponentProps } from 'react-router';
import { Link } from 'react-router-dom';
import { ApplicationState } from '../store';
import * as WeatherForecastsStore from '../store/WeatherForecasts';
import $ from 'jquery';
import { DataManager, Query } from '@syncfusion/ej2-data';
import {
    SpreadsheetComponent,
    SheetsDirective,
    SheetDirective,
    ColumnsDirective,
    RangesDirective,
    RangeDirective,
    RowsDirective,
    RowDirective,
    CellsDirective,
    CellDirective,
    ColumnDirective,
} from "@syncfusion/ej2-react-spreadsheet";
//import { data } from "../Data/DataSource";
import "./sreadsheet.css";
import { createElement } from '@syncfusion/ej2-base'


var sheetn = {
    Rept_Name: "Santxt",
    BoldF: false,
    AeditC: true,
    HideC: false,
    HideR: false,
    Ainsertc: false,
    deleter: false,
    italicc: false,
    Pbackground: '#fff',
    colorc: "rgb(0, 0, 0)",
    Adeletec: false,
    AHideC: false,
    Psheet: false,
    Cunderline: false
};


export class SpreadsheetData extends SpreadsheetComponent {
    state = {
        Rept_Name: "Santxt",
        edit: "true"

    }
    constructor(props) {
        super(props);
        this.boldRight = { fontWeight: "bold", textAlign: "right" };

        this.bold = { fontWeight: "bold" };



    }



    //saveFile() {
    //   var resp = this.spreadsheet.save({ fileName: "Sample" });
    //}

    saveFile() {

        var fl = this.state.Rept_Name;

        this.spreadsheet.saveAsJson().then((response) => {
            var formData = new FormData();
            formData.append('JSONData', JSON.stringify(response.jsonObject.Workbook));
            formData.append('fileName', sheetn.Rept_Name); // 'Santxt');
            formData.append('saveType', 'Xlsx');
            fetch('http://localhost:53142/Home/Save', {
                method: 'POST',
                body: formData
            }).then((response) => {

            });
        });
    }


    //loadFile() {
    //    let request = new XMLHttpRequest();
    //    request.responseType = "blob";
    //    request.onload = () => {
    //        let file = new File([request.response],  this.state.Rept_Name  + ".xlsx");
    //        this.spreadsheet.open({ file: file });
    //    }
    //    request.open("GET", "http://localhost:53142/Files/" +  this.state.Rept_Name  + ".xlsx");
    //    request.send();
    //}
    componentDidMount() {
        let request = new XMLHttpRequest();
        request.responseType = "blob";
        request.onload = () => {
            let file = new File([request.response], "http://localhost:53142/Files/" + sheetn.Rept_Name + ".xlsx");
            this.spreadsheet.open({ file: file });
        }
        request.open("GET", "http://localhost:53142/Files/" + sheetn.Rept_Name + ".xlsx");
        request.send();
        this.spreadsheet.isEdit = false;
    }
       onCreated() {
        this.spreadsheet.cellFormat(
            { fontWeight: "bold", textAlign: "center", verticalAlign: "middle" },
            "A1:F1"

        );
        this.spreadsheet.numberFormat("$#,##0.00", "F2:F31");
        this.spreadsheet.allowEditing = false;

    }

    HtmlConvt() {
        var htmlString = "<table><tbody>";
        var rows = this.spreadsheet.sheets[0].rows;
        for (var i = 0; i < rows.length; i++) {
            htmlString += "<tr>";
            for (var j = 0; j < rows[i].cells.length; j++) {
                htmlString += "<td";
                var cell = rows[i].cells[j];
                if (cell && cell.style) {
                    htmlString += " style='";
                    var style;
                    for (style in cell.style) {
                        switch (style) {
                            case 'fontWeight':
                                htmlString += "font-weight:" + cell.style[style] + ";"
                                break;
                            case 'textAlign':
                                htmlString += "text-align:" + cell.style[style] + ";"
                                break;
                            case 'verticalAlign':
                                htmlString += "vertical-align:" + cell.style[style] + ";"
                                break;
                            case 'backgroundColor':
                                htmlString += "background-color:" + cell.style[style] + ";"
                                break;
                            case 'color':
                                htmlString += "color:" + cell.style[style] + ";"
                                break;
                            case 'fontSize':
                                htmlString += "font-size:" + cell.style[style] + ";"
                                break;
                            case 'fontFamily':
                                htmlString += "font-family:" + cell.style[style] + ";"
                                break;
                        }
                    }
                    htmlString += "'"
                }
                htmlString += (cell && cell.value) ? ">" + cell.value + "</td>" : "></td>";
            }
            htmlString += "</tr>"
        }
        htmlString = "<html><body>" + htmlString + "</tbody></table></body></html>"
        var myBlob = new Blob([htmlString], { type: 'text/html' });
        var anchor = createElement('a', { attrs: { download: "demo.html" } });
        var url = URL.createObjectURL(myBlob);
        anchor.href = url;
        document.body.appendChild(anchor);
        anchor.click();
        URL.revokeObjectURL(url);
        document.body.removeChild(anchor);
    }

    TogelInsertAction() {
        sheetn = {
            ...sheetn, Ainsertc: !sheetn.Ainsertc
        };
        this.spreadsheet.allowInsert = sheetn.Ainsertc;

    }

    TogelHiheColumn() {
        sheetn = {
            ...sheetn, HideC: !sheetn.HideC
        };
        // this.spreadsheet.hideColumn("1");
        //this.spreadsheet.hideRow(1,2,(3);

        this.spreadsheet.hideColumn(3, 4, sheetn.HideC);

    }
    TogelEdit() {
        sheetn = {
            ...sheetn, AeditC: !sheetn.AeditC
        };
        // this.spreadsheet.hideColumn("1");
        //this.spreadsheet.hideRow(1,2,(3);
        this.spreadsheet.allowEditing = sheetn.AeditC;
    }
    handleChangeComplete = (color) => {
        //  sheetn({ ...sheetn, Pbackground: color.hex });
        this.spreadsheet.cellFormat({
            background: "#eee", color: "red"
        }, "A1:C3");

    }
    TogelePSheet() {
        sheetn = {
            ...sheetn, Psheet: !sheetn.Psheet
        };
        // this.spreadsheet.hideColumn("1");
        //this.spreadsheet.hideRow(1,2,(3);
        this.spreadsheet.protectSheet = sheetn.Psheet;
    }
    TogeleDelete() {
        sheetn = {
            ...sheetn, deleter: !sheetn.deleter
        };
        // this.spreadsheet.hideColumn("1");
        //this.spreadsheet.hideRow(1,2,(3);
        this.spreadsheet.allowDelete = sheetn.deleter;
    }

    TogelHiheRow() {
        sheetn = {
            ...sheetn, HideR: !sheetn.HideR
        };
        // this.spreadsheet.hideColumn("1");
        //this.spreadsheet.hideRow(1,2,(3);

        this.spreadsheet.hideRow(2, 2, sheetn.HideR);

    }

    TBoldColumn() {
        sheetn = {
            ...sheetn, BoldF: !sheetn.BoldF
        };

        var bbold = "";
        if (sheetn.BoldF) {
            bbold = "bold";
        }
        this.spreadsheet.cellFormat({
            fontWeight: bbold
        }, "A1:C3");

    }

    TUnderline() {
        sheetn = {
            ...sheetn, Cunderline: !sheetn.Cunderline
        };

        var Cunderline = "";
        if (sheetn.Cunderline) {
            Cunderline = "underline";
        }
        this.spreadsheet.cellFormat({
            textDecoration: Cunderline
        }, "A1:C3");

    }

    boldfont() {
        //this.spreadsheet.setUsedRange.apply("Sanjeeva","1,2");
        //   this.spreadsheet.currentContext.ce
        this.spreadsheet.selectRange("A1:C3");
        this.spreadsheet.cellFormat({
            fontWeight: "bold", textAlign: "Center", textDecoration: "underline"
        }, "A1:C3");
        this.activeSheetIndex = 1;
        //this.spreadsheet.element()
        //this.spreadsheet.cellEditing()
        this.spreadsheet.hideColumn("1");
        this.spreadsheet.hideRow(3);
        this.spreadsheet.allowDelete = false;
        // this.SpreadsheetComponent
        // this.spreadsheet.allowEditing = false;
        this.spreadsheet.allowInsert = false;
        this.spreadsheet.allowCellFormatting = false;
        //this.spreadsheet.se.XLSelection..selectColumns(0, 2);
        //this.XLDragFill.positionAutoFillElement();
        // this.XLDragFill.positionAutoFillElement();
        //this.spreadsheet.selectionSettings.cellFormat(
        //    { fontWeight: "bold",  textAlign: "left", verticalAlign: "middle" }

        //);
    }

    TemplateTrial(fl) {
        sheetn = {
            ...sheetn, Rept_Name: fl
        };

        //this.setState(prevState => {
        //    return {
        //        ...prevState,
        //        Rept_Name: fl
        //    };
        //});
        //  setState({ Rept_Name: fl });
        // this.setState
        let request = new XMLHttpRequest();
        request.responseType = "blob";
        request.onload = () => {
            let file = new File([request.response], "http://localhost:53142/Files/" + sheetn.Rept_Name + ".xlsx");
            this.spreadsheet.open({ file: file });
        }
        request.open("GET", "http://localhost:53142/Files/" + sheetn.Rept_Name + ".xlsx");
        request.send();
    }



    render() {
        return (
            <div>



                <div id="spreadsheet"></div>
                <div className="control-section spreadsheet-control">
                    <div className="left">
                        <button class='e-btn' onClick={this.saveFile.bind(this)}>Save as Excel</button>

                        <ul className="navbar--link">
                            <li className="navbar--link-item" onClick={this.TemplateTrial.bind(this, "santxt")}>Trial Balance</li>
                            <li className="navbar--link-item" onClick={this.TemplateTrial.bind(this, "Template1")}>Ledger</li>
                            <li className="navbar--link-item" onClick={this.TemplateTrial.bind(this, "Template2")}>Report</li>
                        </ul>

                    </div>
                    <div className="center">
                        <SpreadsheetComponent openUrl='http://localhost:53142/Home/Open'
                            saveUrl='http://localhost:53142/Home/Save' ref={(ssObj) => { this.spreadsheet = ssObj; }}
                            cellEdit={false}
                            editSettings={true}>
                            <SheetsDirective>
                                <SheetDirective name='Shipment Details'>
                                    <RangesDirective>
                                        <RangeDirective dataSource={this.data} query={this.query}></RangeDirective>
                                    </RangesDirective>
                                    <ColumnsDirective>                                        <ColumnDirective width={100}></ColumnDirective>
                                        <ColumnDirective width={130}></ColumnDirective>
                                        <ColumnDirective width={150}></ColumnDirective>
                                        <ColumnDirective width={200}></ColumnDirective>
                                        <ColumnDirective width={180}></ColumnDirective>
                                    </ColumnsDirective>
                                </SheetDirective>
                            </SheetsDirective>
                        </SpreadsheetComponent>


                    </div>
                </div>
                <div className="right" >
                    <button class='e-btn' onClick={this.TBoldColumn.bind(this)}>toggle Bold</button>
                    <button class='e-btn' onClick={this.TogelHiheColumn.bind(this)}>Hide/Unhide Column</button>
                    <button class='e-btn' onClick={this.TogelHiheRow.bind(this)}>Hide/Unhide Row</button>
                    <button class='e-btn' onClick={this.TogelEdit.bind(this)}>toggle Edit</button>
                    <button class='e-btn' onClick={this.TogeleDelete.bind(this)}>toggle Delete</button>
                    <button class='e-btn' onClick={this.TogelePSheet.bind(this)}>toggle Protect Sheet</button>

                    <button class='e-btn' onClick={this.handleChangeComplete.bind(sheetn.Pbackground)}>Colour</button>

                    <button class='e-btn' onClick={this.TUnderline.bind(this)}>Toggel Underline</button>
                    <button class='e-btn' onClick={this.TogelInsertAction.bind(this)}>toggle Insert Action</button>
                    <button class='e-btn' onClick={this.HtmlConvt.bind(this)}>To HTML</button>
                    <button class='e-btn' onClick={this.boldfont.bind(this)}>To XML</button>
                </div>
            </div>
        );
    }
}
