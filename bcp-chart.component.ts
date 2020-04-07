import { Component, OnInit } from '@angular/core';
import { BcpChartService } from '../providers/bcp-chart.service';
import { ActivatedRoute, Router } from '@angular/router';
import { BCPDailyUpdate } from '../models/BCPDailyUpdate';
import { BcpAssociateTrackerService } from '../providers/bcp-associates-tracker.service';
import * as FileSaver from 'file-saver';
import * as XLSX from 'xlsx';
import * as moment from 'moment';
import { BCPDetailsUpdate } from '../models/BCPDetailsUpdate';

const EXCEL_TYPE = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=UTF-8';
const EXCEL_EXTENSION = '.xlsx';

declare var require: any;
@Component({
    selector: 'app-bcp-chart',
    templateUrl: './bcp-chart.component.html',
    styleUrls: ['./bcp-chart.component.scss']
})
export class BcpChartComponent implements OnInit {
    constructor(private bcpChartService: BcpChartService,
        private bcpAssociateTrackerService: BcpAssociateTrackerService,
        private router: Router,
        private route: ActivatedRoute) { }

    keyword = 'name';
    projectId: any;
    availableDate: any = [];
    attendanceData: any = [];
    TotalAttendanceDownloadData: any = [];
    deviceType: any = [];
    accountCount: any;
    protocolData: any = [];
    wfhData: any = [];
    piiAccessData: any = [];
    byodData: any = [];

    ngOnInit() {
        this.route.params.subscribe(params => { this.projectId = params["id"] });
        this.bcpChartService.getBCPDataTrackerHistoryCount(this.projectId).subscribe(data => {
            this.accountCount = data;
            this.getBcpDetailsUpdateData(this.projectId);
        });
    }

    downloadProtocolReport() {
        if (this.protocolData.length > 0) {
            this.exportExcel(this.protocolData, this.projectId + "Protocol", "ProtocolReport");
        }
    }

    exportExcel(json: any[], fileName: string, sheetName: string) {
        var wb = { SheetNames: [], Sheets: {} };
        const worksheet1: XLSX.WorkSheet = XLSX.utils.json_to_sheet(json);
        wb.SheetNames.push("ProtocolReport");
        wb.Sheets["ProtocolReport"] = worksheet1;
        const excelBuffer: any = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
        const data: Blob = new Blob([excelBuffer], { type: EXCEL_TYPE });
        FileSaver.saveAs(data, fileName + '_export_' + EXCEL_EXTENSION);
    }

    NavigateToUserTracker() {
        this.router.navigate(['/bcm-user-tracker', this.projectId]);
    }

    stringToDate(dateString: string) {
        var dateParts = dateString.split('-');
        return new Date(parseInt(dateParts[2]), parseInt(dateParts[1]) - 1, parseInt(dateParts[0]));
    }

    fillMissingDates(actualDatesinDb: any[]) {
        let datesAfterInsertingHolidays = [];

        for (let i = 0; i < actualDatesinDb.length - 1; i++) {
            var start = moment(this.stringToDate(actualDatesinDb[i]));
            var end = moment(this.stringToDate(actualDatesinDb[i + 1]));
            var diff = Math.abs(end.diff(start, 'days'));
            if (diff > 1) {
                if (new Date(start.toLocaleString()).getDay() !== 6 && new Date(start.toLocaleString()).getDay() !== 0) {
                    datesAfterInsertingHolidays.push(start.format("DD-MM-YYYY"));
                }
                for (let j = 1; j <= diff; j++) {
                    var missingDate = moment(start, "DD-MM-YYYY").add('days', j).toLocaleString();
                    if (new Date(missingDate).getDay() !== 6 && new Date(missingDate).getDay() !== 0) {
                        datesAfterInsertingHolidays.push(moment(new Date(missingDate)).format("DD-MM-YYYY"));
                    }
                    diff--;
                }
            } else {
                if (new Date(start.toLocaleString()).getDay() !== 6 && new Date(start.toLocaleString()).getDay() !== 0) {
                    datesAfterInsertingHolidays.push(start.format("DD-MM-YYYY"));
                }
            }
        }
        return datesAfterInsertingHolidays;
    }

    getBcpDetailsUpdateData(projectId) {
        this.bcpChartService.getBCPDataTrackerHistory(projectId).subscribe(data => {
            this.bcpAssociateTrackerService.getBcpAssociateTracker(projectId).subscribe(model => {
                let chartData = [];
                data.bcpDetailsUpdate.forEach(bcpDetails => {
                    var bcpModelDetails = model.userDetail.filter(x => x.AssociateId == bcpDetails.AssociateID);
                    var modelDetails = bcpModelDetails.map(userDetails => ({
                        AccountID: userDetails.AccountID,
                        AccountName: userDetails.AccountName,
                        AssociateID: userDetails.AssociateId,
                        AssociateName: userDetails.AssociateName,
                        CurrentEnabledforWFH: bcpDetails.CurrentEnabledforWFH,
                        WFHDeviceType: bcpDetails.WFHDeviceType,
                        Comments: bcpDetails.Comments,
                        PersonalReason: bcpDetails.PersonalReason,
                        AssetId: bcpDetails.AssetId,
                        PIIDataAccess: bcpDetails.PIIDataAccess,
                        Protocol: bcpDetails.Protocol,
                        BYODCompliance: bcpDetails.BYODCompliance,
                        Dongles: bcpDetails.Dongles,
                        UpdateDate: bcpDetails.UpdateDate,
                        UniqueId: bcpDetails.UniqueId
                    }));
                    chartData.push(modelDetails);
                });
                console.log(chartData);
                this.getChartData(chartData);
            });
        });
        this.bcpChartService.getAccountAttendanceData(projectId).subscribe((response: BCPDailyUpdate[]) => {
            const uniqueUpdateDate = this.fillMissingDates([...new Set(response.map(item => item.UpdateDate))]);
            uniqueUpdateDate.forEach((updateDate: any) => {
                const uniqueYes = response.filter(item => item.UpdateDate == updateDate && item.Attendance == "No");
                const uniqueYesCount = this.accountCount - uniqueYes.length;
                const percent = (uniqueYesCount / this.accountCount) * 100;
                const roundPer = parseFloat(percent.toString()).toFixed(2);
                this.attendanceData.push({ date: updateDate, count: +roundPer });
            });
            this.attendanceGraph(this.attendanceData);
        });
    }

    getChartData(chartData) {
        console.log(chartData);
        this.getWFHReadiness(chartData);
        this.getDeviceType(chartData);
        this.getPersonalReason(chartData);
        this.getProtocolType(chartData);
        this.getPiiAcess(chartData);
        this.getBYODCompliance(chartData);
    }

    private getWFHReadiness(chartData) {
        debugger;
        var wfhRedinessYes;
        var wfhRedinessNo;
        var uniqueYes = chartData.filter(item => item.CurrentEnabledforWFH == "Yes");
        uniqueYes.forEach(x => {
            this.wfhData.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, CurrentEnabledforWFH: "Yes" });
        });
        var uniqueNo = chartData.filter(item => item.CurrentEnabledforWFH == "No");
        uniqueNo.forEach(x => {
            this.wfhData.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, CurrentEnabledforWFH: "No" });
        });
        const uniqueNoCount = chartData.length - uniqueYes.length;
        wfhRedinessYes = parseFloat(((uniqueYes.length / chartData.length) * 100).toFixed(2));
        wfhRedinessNo = parseFloat(((uniqueNoCount / chartData.length) * 100).toFixed(2));
        this.workFromHomeGraph(wfhRedinessYes, wfhRedinessNo);
    }

    WFHReadinessExcelSheetData() {
        console.log(this.wfhData);
        if (this.wfhData.length > 0) {
            this.exportExcel(this.wfhData, this.projectId + "WFHReadiness", "WFHReadinessDetails");
        }
    }

    private getPiiAcess(chartData) {
        debugger
        var PIIDataAccessYes;
        var PIIDataAccessNo;
        var uniqueYes = chartData.filter(item => item.PIIDataAccess == "Yes");
        uniqueYes.forEach(x => {
            this.piiAccessData.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, PIIDataAccess: "Yes" });
        });
        var uniqueNo = chartData.filter(item => item.PIIDataAccess == "No");
        uniqueNo.forEach(x => {
            this.piiAccessData.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, PIIDataAccess: "No" });
        });
        const uniqueNoCount = chartData.length - uniqueYes.length;
        PIIDataAccessYes = parseFloat(((uniqueYes.length / chartData.length) * 100).toFixed(2));
        PIIDataAccessNo = parseFloat(((uniqueNoCount / chartData.length) * 100).toFixed(2));
        this.piiAccessGraph(PIIDataAccessYes, PIIDataAccessNo);
    }

    PiiAcessExcelSheetData() {
        console.log(this.piiAccessData);
        if (this.piiAccessData.length > 0) {
            this.exportExcel(this.piiAccessData, this.projectId + "PIIDataAccess", "PIIDataAccessDetails");

        }
    }

    private getBYODCompliance(chartData) {
        debugger
        var BYODComplianceYes;
        var BYODComplianceNo;
        var uniqueYes = chartData.filter(item => item.BYODCompliance == "Yes");
        uniqueYes.forEach(x => {
            this.byodData.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, BYODCompliance: "Yes" });
        });
        var uniqueNo = chartData.filter(item => item.BYODCompliance == "No");
        uniqueNo.forEach(x => {
            this.byodData.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, BYODCompliance: "No" });
        });
        const uniqueNoCount = chartData.length - uniqueYes.length;
        BYODComplianceYes = parseFloat(((uniqueYes.length / chartData.length) * 100).toFixed(2));
        BYODComplianceNo = parseFloat(((uniqueNoCount / chartData.length) * 100).toFixed(2));
        this.BYODComplianceGraph(BYODComplianceYes, BYODComplianceNo);
    }

    BYODComplianceExcelSheetData() {
        console.log(this.byodData);
        if (this.byodData.length > 0) {
            this.exportExcel(this.byodData, this.projectId + "BYODCompliance", "BYODComplianceDetails");
        }
    }

    private getDeviceType(chartData) {
        var personalDevice = [];
        var cognizantDevice = [];
        var customerDevice = [];
        var cognizantBYODs = [];
        this.deviceType = [];

        var personaltemp = chartData.filter(item => item.WFHDeviceType == "Personal Device");
        personaltemp.forEach(x => {
            this.deviceType.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, DeviceType: x.WFHDeviceType });
        });
        personalDevice.push({ count: personaltemp.length, data: personaltemp });
        var cognizantDevicetemp = chartData.filter(item => item.WFHDeviceType == "Cognizant Device");
        cognizantDevicetemp.forEach(x => {
            this.deviceType.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, DeviceType: x.WFHDeviceType });
        });
        cognizantDevice.push({ count: cognizantDevicetemp.length, data: personaltemp });
        var customerDevicetemp = chartData.filter(item => item.WFHDeviceType == "Customer Device");
        customerDevicetemp.forEach(x => {
            this.deviceType.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, DeviceType: x.WFHDeviceType });
        });
        customerDevice.push({ count: customerDevicetemp.length, data: personaltemp });
        var cognizantBOYDstemp = chartData.filter(item => item.WFHDeviceType == "Cognizant BYOD");
        cognizantBOYDstemp.forEach(x => {
            this.deviceType.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, DeviceType: x.WFHDeviceType });
        });
        cognizantBYODs.push({ count: cognizantBOYDstemp.length, data: personaltemp });

        this.deviceTypeGraph(personalDevice, cognizantDevice, customerDevice, cognizantBYODs);
    }

    private getPersonalReason(chartData) {
        var noDevice = [];
        var unplannedLeave = [];
        var plannedLeave = [];
        var workingAtOffice = [];
        var connectivity = [];
        var covid19 = [];

        var nodevice = chartData.filter(item => item.PersonalReason == "No device");
        nodevice.forEach(x => {
            this.availableDate.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, PersonalLeave: "No device" });
        });
        noDevice.push({ count: nodevice.length, data: nodevice });
        var unplanned = chartData.filter(item => item.PersonalReason == "unplanned leave");
        unplanned.forEach(x => {
            this.availableDate.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, PersonalLeave: "Unplanned leave" });
        });
        unplannedLeave.push({ count: unplanned.length, data: unplanned });
        var planned = chartData.filter(item => item.PersonalReason == "planned leave");
        planned.forEach(x => {
            this.availableDate.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, PersonalLeave: "Planned leave" });
        });
        plannedLeave.push({ count: planned.length, data: planned });
        var workAtOffice = chartData.filter(item => item.PersonalReason == "working at office");
        workAtOffice.forEach(x => {
            this.availableDate.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, PersonalLeave: "Working at office" });
        });
        workingAtOffice.push({ count: workAtOffice.length, data: workAtOffice });
        var connect = chartData.filter(item => item.PersonalReason == "Connectivity");
        connect.forEach(x => {
            this.availableDate.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, PersonalLeave: "Connectivity" });
        });
        connectivity.push({ count: connect.length, data: connect });
        var covid = chartData.filter(item => item.PersonalReason == "COVID19");
        covid.forEach(x => {
            this.availableDate.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, PersonalLeave: "COVID19" });
        });
        covid19.push({ count: covid.length, data: covid });

        this.personalReasonGraph(noDevice, unplannedLeave, plannedLeave, workingAtOffice, connectivity, covid19);
    }

    private getProtocolType(chartData) {
        var protocolA = chartData.filter(item => item.Protocol == "Protocol A");

        protocolA.forEach(x => {
            this.protocolData.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, Protocol: "Protocol A" });
        });

        var protocolB1 = chartData.filter(item => item.Protocol == "Protocol B.1");

        protocolB1.forEach(x => {
            this.protocolData.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, Protocol: "Protocol B.1" });
        });

        var protocolB2 = chartData.filter(item => item.Protocol == "Protocol B.2");

        protocolB2.forEach(x => {
            this.protocolData.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, Protocol: "Protocol B.2" });
        });

        var protocolB3 = chartData.filter(item => item.Protocol == "Protocol B.3");

        protocolB3.forEach(x => {
            this.protocolData.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, Protocol: "Protocol B.3" });
        });

        var protocolB4 = chartData.filter(item => item.Protocol == "Protocol B.4");

        protocolB4.forEach(x => {
            this.protocolData.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, Protocol: "Protocol B.4" });
        });

        var protocolC1 = chartData.filter(item => item.Protocol == "Protocol C.1");

        protocolC1.forEach(x => {
            this.protocolData.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, Protocol: "Protocol C.1" });
        });

        var protocolC2 = chartData.filter(item => item.Protocol == "Protocol C.2");

        protocolC2.forEach(x => {
            this.protocolData.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, Protocol: "Protocol C.2" });
        });

        var protocolC3 = chartData.filter(item => item.Protocol == "Protocol C.3");

        protocolC3.forEach(x => {
            this.protocolData.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, Protocol: "Protocol C.3" });
        });

        var protocolC4 = chartData.filter(item => item.Protocol == "Protocol C.4");

        protocolC4.forEach(x => {
            this.protocolData.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, Protocol: "Protocol C.4" });
        });

        var protocolD = chartData.filter(item => item.Protocol == "Protocol D");

        protocolD.forEach(x => {
            this.protocolData.push({ AccountID: x.AccountId, AccountName: x.AccountName, AssociateId: x.AssociateID, AssociateName: x.AssociateName, Date: x.UpdateDate, Protocol: "Protocol D" });
        });

        this.protocolGraph(chartData, protocolA.length, protocolB1.length, protocolB2.length, protocolB3.length,
            protocolB4.length, protocolC1.length, protocolC2.length, protocolC3.length, protocolC4.length, protocolD.length);
    }

    exportContainer1() {
        if (this.deviceType.length > 0) {
            this.exportExcel(this.deviceType, this.projectId + "DeviceType", "DeviceType");
        }
    }

    deviceTypeGraph(personalDevice, cognizantDevice, customerDevice, cognizantBYODs) {
        var Yaxis_personal = [];
        var Yaxis_cts = [];
        var Yaxis_cus = [];
        var Yaxis_byod = [];
        var fileName = "";
        var sheetName = "Device Availability";
        var resultGraph = [];
        personalDevice.forEach(element => {
            Yaxis_personal.push(element.count);
        });
        cognizantDevice.forEach(element => {
            Yaxis_cts.push(element.count);
        });
        customerDevice.forEach(element => {
            Yaxis_cus.push(element.count);
        });
        cognizantBYODs.forEach(element => {
            Yaxis_byod.push(element.count);
        });

        var Highcharts = require('highcharts');
        require('highcharts/modules/exporting')(Highcharts);
        Highcharts.chart('container1', {
            chart: {
                type: 'column'
            },
            title: {
                text: 'Personal/Customer/Cognizant Device Availability'
            },
            xAxis: {
                categories: ['Device']
            },
            yAxis: {
                title: {
                    text: ' Count '
                }
            },
            credits: {
                enabled: false
            },
            plotOptions: {
                series: {
                    cursor: 'pointer',
                    point: {
                        events: {
                            click: function () {
                                if (this.series.name == "Cognizant BYOD") {
                                    resultGraph = cognizantBYODs[0].data;
                                    fileName = "BYOD";
                                } else if (this.series.name == "Cognizant Device") {

                                    resultGraph = cognizantDevice[0].data;
                                    fileName = "Cognizant";
                                } else if (this.series.name == "Customer Device") {
                                    resultGraph = customerDevice[0].data;
                                    fileName = "Customer";
                                } else if (this.series.name == "Personal Device") {
                                    resultGraph = personalDevice[0].data;
                                    fileName = "Personal";
                                }
                                var wb = { SheetNames: [], Sheets: {} };
                                const worksheet1: XLSX.WorkSheet = XLSX.utils.json_to_sheet(resultGraph);
                                wb.SheetNames.push(sheetName);
                                wb.Sheets[sheetName] = worksheet1;
                                const excelBuffer: any = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
                                const data: Blob = new Blob([excelBuffer], { type: EXCEL_TYPE });
                                FileSaver.saveAs(data, fileName + '_export_' + EXCEL_EXTENSION);
                            }
                        }
                    }
                },
                line: {
                    dataLabels: {
                        enabled: true
                    },
                    enableMouseTracking: false
                }
            },
            series: [{
                name: 'Personal Device',
                data: Yaxis_personal
            },
            {
                name: 'Cognizant Device',
                data: Yaxis_cts
            },
            {
                name: 'Customer Device',
                data: Yaxis_cus
            },
            {
                name: 'Cognizant BYOD',
                data: Yaxis_byod
            }]
        });
    }

    piiAccessGraph(PIIDataAccessYes, PIIDataAccessNo) {
        var Highcharts = require('highcharts');
        require('highcharts/modules/exporting')(Highcharts);
        Highcharts.chart('container6', {
            chart: {
                plotBackgroundColor: null,
                plotBorderWidth: null,
                plotShadow: false,
                type: 'pie'
            },
            title: {
                text: 'PII Access'
            },
            tooltip: {
                pointFormat: '{series.name}: <b>{point.y}%</b>'
            },
            accessibility: {
                point: {
                    valueSuffix: '%'
                }
            },
            credits: {
                enabled: false
            },
            plotOptions: {
                pie: {
                    allowPointSelect: true,
                    cursor: 'pointer',
                    dataLabels: {
                        enabled: true,
                        format: '<b>{point.name}</b>: {point.y} %'
                    }
                },
            },
            series: [{
                name: 'PII Access',
                colorByPoint: true,
                data: [{
                    name: 'PII Access - yes',
                    y: PIIDataAccessYes,
                    sliced: true,
                    selected: true
                }, {
                    name: 'PII Access - No',
                    y: PIIDataAccessNo
                }]
            }]
        });
    }
    workFromHomeGraph(wfhRedinessYes, wfhRedinessNo) {
        var Highcharts = require('highcharts');
        require('highcharts/modules/exporting')(Highcharts);
        Highcharts.chart('container2', {
            chart: {
                plotBackgroundColor: null,
                plotBorderWidth: null,
                plotShadow: false,
                type: 'pie'
            },
            title: {
                text: 'Work From Home'
            },
            tooltip: {
                pointFormat: '{series.name}: <b>{point.percentage:.1f}%</b>'
            },
            accessibility: {
                point: {
                    valueSuffix: '%'
                }
            },
            credits: {
                enabled: false
            },
            plotOptions: {
                pie: {
                    allowPointSelect: true,
                    cursor: 'pointer',
                    dataLabels: {
                        enabled: true,
                        format: '<b>{point.name}</b>: {point.percentage:.1f} %'
                    }
                }
            },
            series: [{
                name: 'Work From Home',
                colorByPoint: true,
                data: [{
                    name: 'WFH - yes',
                    y: wfhRedinessYes,
                    sliced: true,
                    selected: true
                }, {
                    name: 'WFH - No',
                    y: wfhRedinessNo
                }]
            }]
        });
    }
    BYODComplianceGraph(BYODComplianceYes, BYODComplianceNo) {
        var Highcharts = require('highcharts');
        require('highcharts/modules/exporting')(Highcharts);
        Highcharts.chart('container7', {
            chart: {
                plotBackgroundColor: null,
                plotBorderWidth: null,
                plotShadow: false,
                type: 'pie'
            },
            title: {
                text: 'BYODCompliance'
            },
            tooltip: {
                pointFormat: '{series.name}: <b>{point.percentage:.1f}%</b>'
            },
            accessibility: {
                point: {
                    valueSuffix: '%'
                }
            },
            credits: {
                enabled: false
            },
            plotOptions: {
                pie: {
                    allowPointSelect: true,
                    cursor: 'pointer',
                    dataLabels: {
                        enabled: true,
                        format: '<b>{point.name}</b>: {point.percentage:.1f} %'
                    }
                }
            },
            series: [{
                name: 'BYODCompliance',
                colorByPoint: true,
                data: [{
                    name: 'BYODCompliance - yes',
                    y: BYODComplianceYes,
                    sliced: true,
                    selected: true
                }, {
                    name: 'BYODCompliance - No',
                    y: BYODComplianceNo
                }]
            }]
        });
    }

    personalLeaveExcelSheetData() {
        if (this.availableDate.length > 0) {
            this.exportExcel(this.availableDate, this.projectId + "PersonalReason", "PersonalReason");
        }
    }

    personalReasonGraph(noDevice, unplannedLeave, plannedLeave, workingAtOffice, connectivity, covid19) {
        var noDeviceYaxis = [];
        var unplannedLeaveYaxis = [];
        var plannedLeaveYaxis = [];
        var workingAtOfficeYaxis = [];
        var connectivityYaxis = [];
        var covidYaxis = [];
        var fileName = "";
        var sheetName = "Personal Reason";
        var resultGraph = [];

        noDevice.forEach(element => {
            noDeviceYaxis.push(element.count);
        });
        unplannedLeave.forEach(element => {
            unplannedLeaveYaxis.push(element.count);
        });
        plannedLeave.forEach(element => {
            plannedLeaveYaxis.push(element.count);
        });
        workingAtOffice.forEach(element => {
            workingAtOfficeYaxis.push(element.count);
        });
        connectivity.forEach(element => {
            connectivityYaxis.push(element.count);
        });
        covid19.forEach(element => {
            covidYaxis.push(element.count);
        });
        var Highcharts = require('highcharts');
        require('highcharts/modules/exporting')(Highcharts);
        Highcharts.chart('container3', {
            chart: {
                type: 'column'
            },
            title: {
                text: ' Personal Reason'
            },
            xAxis: {
                categories: ['Reason']
            },
            yAxis: {
                title: {
                    text: ' Count '
                }
            },
            credits: {
                enabled: false
            },
            plotOptions: {
                line: {
                    dataLabels: {
                        enabled: true
                    },
                    enableMouseTracking: false
                },
                series: {
                    cursor: 'pointer',
                    point: {
                        events: {
                            click: function (event) {
                                if (this.series.name == "No device") {
                                    resultGraph = noDevice[0].data.map(item => ({
                                        AccountID: item.Title,
                                        AssociateID: item.AssociateID,
                                        PersonalLeave: item.PersonalLeave
                                    }));
                                    fileName = "Nodevice";
                                }
                                else if (this.series.name == "COVID19") {
                                    resultGraph = covid19[0].data.map(item => ({
                                        AccountID: item.Title,
                                        AssociateID: item.AssociateID,
                                        PersonalLeave: item.PersonalLeave
                                    }));
                                    fileName = "COVID19";
                                }
                                else if (this.series.name == "Unplanned leave") {
                                    resultGraph = unplannedLeave[0].data.map(item => ({
                                        AccountID: item.Title,
                                        AssociateID: item.AssociateID,
                                        PersonalLeave: item.PersonalLeave
                                    }));
                                    fileName = "Unplanned";
                                }
                                else if (this.series.name == "Planned leave") {
                                    resultGraph = plannedLeave[0].data.map(item => ({
                                        AccountID: item.Title,
                                        AssociateID: item.AssociateID,
                                        PersonalLeave: item.PersonalLeave
                                    }));
                                    fileName = "Planned";
                                }
                                else if (this.series.name == "Working at office") {
                                    resultGraph = workingAtOffice[0].data.map(item => ({
                                        AccountID: item.Title,
                                        AssociateID: item.AssociateID,
                                        PersonalLeave: item.PersonalLeave
                                    }));
                                    fileName = "WAO";
                                }
                                else if (this.series.name == "Connectivity") {
                                    resultGraph = connectivity[0].data.map(item => ({
                                        AccountID: item.Title,
                                        AssociateID: item.AssociateID,
                                        PersonalLeave: item.PersonalLeave
                                    }));
                                    fileName = "Connectivity";
                                }

                                var wb = { SheetNames: [], Sheets: {} };
                                const worksheet1: XLSX.WorkSheet = XLSX.utils.json_to_sheet(resultGraph);
                                wb.SheetNames.push(sheetName);
                                wb.Sheets[sheetName] = worksheet1;
                                const excelBuffer: any = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
                                const data: Blob = new Blob([excelBuffer], { type: EXCEL_TYPE });
                                FileSaver.saveAs(data, fileName + '_export_' + EXCEL_EXTENSION);
                            }
                        }
                    }
                }
            },
            series: [{
                name: 'No device',
                data: noDeviceYaxis
            },
            {
                name: 'Unplanned leave',
                data: unplannedLeaveYaxis
            },
            {
                name: 'Planned leave',
                data: plannedLeaveYaxis
            },
            {
                name: 'Working at office',
                data: workingAtOfficeYaxis
            },
            {
                name: 'Connectivity',
                data: connectivityYaxis
            },
            {
                name: 'COVID19',
                data: covidYaxis
            }]
        });
    }

    attendanceGraph(attendance) {
        var xaxis = [];
        var yaxis = [];
        attendance.forEach(element => {
            xaxis.push(element.date);
            yaxis.push(element.count);
        });
        var Highcharts = require('highcharts');
        require('highcharts/modules/exporting')(Highcharts);
        Highcharts.chart('container4', {
            chart: {
                type: 'column'
            },
            title: {
                text: 'Attendance'
            },
            xAxis: {
                categories: xaxis
            },
            yAxis: {
                title: {
                    text: ' Percentage '
                }
            },
            credits: {
                enabled: false
            },
            plotOptions: {
                series: {
                    point: {
                        events: {
                            click: (currentValue) => {
                                this.attendanceDownloadByDate(currentValue.point.category);
                            }
                        }
                    }
                }
            },
            series: [{
                name: 'Attendance',
                data: yaxis
            }]
        });
    }

    attendanceDownloadByDate(specificDate?: any) {
        let PresentMembers = [];
        let AbsentMembers = [];
        this.TotalAttendanceDownloadData = [];
        this.bcpAssociateTrackerService.getBcpAssociateTracker(this.projectId).subscribe(model => {
            let totalAccountMembers = model.userDetail;
            this.bcpChartService.getAccountAttendanceData(this.projectId).subscribe((response: BCPDailyUpdate[]) => {
                let uniqueUpdateDate;
                if (specificDate) {
                    uniqueUpdateDate = [specificDate];
                } else {
                    uniqueUpdateDate = this.fillMissingDates([...new Set(response.map(item => item.UpdateDate))]);
                }
                uniqueUpdateDate.forEach(selectedDate => {
                    let membersAbsentOnSelectedDate = response.filter(item => item.UpdateDate == selectedDate && item.Attendance == "No");
                    PresentMembers = totalAccountMembers.filter(x => !membersAbsentOnSelectedDate.map(y => y.AssociateID).includes(x.AssociateId))
                        .map(a => ({
                            AccountID: a.AccountID,
                            AccountName: a.AccountName,
                            AssociateId: a.AssociateId,
                            AssociateName: a.AssociateName,
                            Date: selectedDate,
                            Attendance: "Yes"
                        }));

                    let membersPresentOnSelectedDate = response.filter(item => item.UpdateDate == selectedDate && item.Attendance == "No");
                    AbsentMembers = totalAccountMembers.filter(x => membersPresentOnSelectedDate.map(y => y.AssociateID).includes(x.AssociateId))
                        .map(a => ({
                            AccountID: a.AccountID,
                            AccountName: a.AccountName,
                            AssociateId: a.AssociateId,
                            AssociateName: a.AssociateName,
                            Date: selectedDate,
                            Attendance: "No"
                        }));

                    this.TotalAttendanceDownloadData.push.apply(this.TotalAttendanceDownloadData, AbsentMembers.concat(PresentMembers))

                });

                if (this.TotalAttendanceDownloadData.length > 0) {
                    this.exportExcel(this.TotalAttendanceDownloadData, this.projectId + "Attendance", "AttendanceDetails");

                }
            });
        });
    }

    protocolGraph(chartData, protocolA, protocolB1, protocolB2, protocolB3, protocolB4, protocolC1, protocolC2, protocolC3, protocolC4, protocolD) {

        var fileName = "";
        var sheetName = "Protocol Details";
        var resultGraph = [];

        var Highcharts = require('highcharts');
        require('highcharts/modules/exporting')(Highcharts);
        Highcharts.chart('container5', {
            chart: {
                type: 'bar'
            },
            title: {
                text: 'Protocol'
            },
            xAxis: {
                categories: [
                    'Protocol A', 'Protocol B.1', 'Protocol B.2', 'Protocol B.3', 'Protocol B.4',
                    'Protocol C.1', 'Protocol C.2', 'Protocol C.3', 'Protocol C.4', 'Protocol D'
                ],
                title: {
                    text: null
                }
            },
            yAxis: {
                min: 0,
                title: {
                    text: 'Count'
                }
            },
            credits: {
                enabled: false
            },
            plotOptions: {
                series: {
                    cursor: 'pointer',
                    point: {
                        events: {
                            click: function () {
                                var protocol;
                                if (this.x == '0') {
                                    protocol = "A";
                                }
                                else if (this.x == '1') {
                                    protocol = "B1";
                                }
                                else if (this.x == '2') {
                                    protocol = "B2";
                                }
                                else if (this.x == '3') {
                                    protocol = "B3";
                                }
                                else if (this.x == '4') {
                                    protocol = "B4";
                                }
                                else if (this.x == '5') {
                                    protocol = "C1";
                                }
                                else if (this.x == '6') {
                                    protocol = "C2";
                                }
                                else if (this.x == '7') {
                                    protocol = "C3";
                                }
                                else if (this.x == '8') {
                                    protocol = "C4";
                                }
                                else if (this.x == '9') {
                                    protocol = "D";
                                }


                                if (protocolA == this.y) {
                                    resultGraph = chartData.filter(x => x.Protocol == "ProtocolA").map(item => {
                                        return new (
                                            item.Title,
                                            item.AssociateID,
                                            item.Protocol
                                        )
                                    });
                                    fileName = "ProtocolA";
                                }

                                else if (protocolB1 == this.y) {
                                    resultGraph = chartData.filter(x => x.Protocol == "ProtocolB1").map(item => {
                                        return new (
                                            item.Title,
                                            item.AssociateID,
                                            item.Protocol
                                        )
                                    });
                                    fileName = "ProtocolB1";
                                }

                                else if (protocolB2 == this.y) {
                                    resultGraph = chartData.filter(x => x.Protocol == "ProtocolB2").map(item => {
                                        return new (
                                            item.Title,
                                            item.AssociateID,
                                            item.Protocol
                                        )
                                    });
                                    fileName = "ProtocolB2";
                                }

                                else if (protocolB3 == this.y) {
                                    resultGraph = chartData.filter(x => x.Protocol == "ProtocolB3").map(item => {
                                        return new (
                                            item.Title,
                                            item.AssociateID,
                                            item.Protocol
                                        )
                                    });
                                    fileName = "ProtocolB3";
                                }

                                else if (protocolB4 == this.y) {
                                    resultGraph = chartData.filter(x => x.Protocol == "ProtocolB4").map(item => {
                                        return new (
                                            item.Title,
                                            item.AssociateID,
                                            item.Protocol
                                        )
                                    });
                                    fileName = "ProtocolB4";
                                }

                                else if (protocolC1 == this.y) {
                                    resultGraph = chartData.filter(x => x.Protocol == "ProtocolC1").map(item => {
                                        return new (
                                            item.Title,
                                            item.AssociateID,
                                            item.Protocol
                                        )
                                    });
                                    fileName = "ProtocolC1";
                                }

                                else if (protocolC2 == this.y) {
                                    resultGraph = chartData.filter(x => x.Protocol == "ProtocolC2").map(item => {
                                        return new (
                                            item.Title,
                                            item.AssociateID,
                                            item.Protocol
                                        )
                                    });
                                    fileName = "ProtocolC2";
                                }

                                else if (protocolC3 == this.y) {
                                    resultGraph = chartData.filter(x => x.Protocol == "ProtocolC3").map(item => {
                                        return new (
                                            item.Title,
                                            item.AssociateID,
                                            item.Protocol
                                        )
                                    });
                                    fileName = "ProtocolC3";
                                }

                                else if (protocolC4 == this.y) {
                                    resultGraph = chartData.filter(x => x.Protocol == "ProtocolC4").map(item => {
                                        return new (
                                            item.Title,
                                            item.AssociateID,
                                            item.Protocol
                                        )
                                    });
                                    fileName = "ProtocolC4";
                                }

                                else if (protocolD == this.y) {
                                    resultGraph = chartData.filter(x => x.Protocol == "ProtocolD").map(item => {
                                        return new (
                                            item.Title,
                                            item.AssociateID,
                                            item.Protocol
                                        )
                                    });
                                    fileName = "ProtocolD";
                                }

                                var wb = { SheetNames: [], Sheets: {} };
                                const worksheet1: XLSX.WorkSheet = XLSX.utils.json_to_sheet(resultGraph);
                                wb.SheetNames.push(sheetName);
                                wb.Sheets[sheetName] = worksheet1;
                                const excelBuffer: any = XLSX.write(wb, { bookType: 'xlsx', type: 'array' });
                                const data: Blob = new Blob([excelBuffer], { type: EXCEL_TYPE });
                                FileSaver.saveAs(data, fileName + '_export_' + EXCEL_EXTENSION);
                            }
                        }
                    }
                },
                line: {
                    dataLabels: {
                        enabled: true
                    },
                    enableMouseTracking: false
                }
            },
            series: [{
                name: 'Protocol Count',
                data: [protocolA, protocolB1, protocolB2, protocolB3, protocolB4, protocolC1, protocolC2, protocolC3,
                    protocolC4, protocolD]
            }]
        });
    }
}