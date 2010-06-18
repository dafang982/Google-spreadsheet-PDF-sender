// Copyright 2010 Jiayao Yu
//
// Licensed under the Apache License, Version 2.0 (the "License");
// you may not use this file except in compliance with the License.
// You may obtain a copy of the License at
//
// http://www.apache.org/licenses/LICENSE-2.0
//
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.

// Google spreadsheet script for sending spreadsheet as PDF

var ENABLED_CELL = 1;
var TOKEN_CELL = 2;
var EMAIL_CELL = 3;
var BCC_CELL = 4;
var SUBJECT_CELL = 5;
var BODY_CELL = 6;
var SHEET_NAME_CELL = 7;
var SHEET_GID_CELL = 8;

var SPREADSHEET_URL = "http://spreadsheets.google.com/feeds/download/spreadsheets/Export?key=";
var MAX_CONFIG_ROWS = 1000;
var MAX_EXPORT_SHEETS = 50;
var EXPORT_FORMAT = "pdf";

function onOpen() {
    var ss = SpreadsheetApp.getActiveSpreadsheet(),
	menuEntries = [ {name: "Send as PDF", functionName: "sendAsPdf"}];
    ss.addMenu("PDF Sender", menuEntries);
    var configSheet = getConfigSheet();
    configSheet.getRange(1, ENABLED_CELL).setValue("Enabled");
    configSheet.getRange(1, TOKEN_CELL).setValue("Auth Token");
    configSheet.getRange(1, EMAIL_CELL).setValue("Email");
    configSheet.getRange(1, BCC_CELL).setValue("Bcc");
    configSheet.getRange(1, SUBJECT_CELL).setValue("Subject");
    configSheet.getRange(1, BODY_CELL).setValue("Body");
    configSheet.getRange(1, SHEET_NAME_CELL).setValue("Export sheet name");
    configSheet.getRange(1, SHEET_GID_CELL).setValue("Export sheet gid");
}

function sendAsPdf() {
    var configSheet = getConfigSheet();
    for (var i = 2; i < MAX_CONFIG_ROWS; i++) {
	if (configSheet.getRange(i, 2).getValue()) {
	    sendForConfigRow(i);
	} else {
	    break;
	}
    }
}

function sendForConfigRow(row) {
    if (getConfig(row, ENABLED_CELL) != true) {
	return;
    }
    var attachments = [];
    var sheetName = getConfig(row, SHEET_NAME_CELL);
    for (var i = 0; i < MAX_EXPORT_SHEETS; i++) {
	var sheetGid = getConfig(row, SHEET_GID_CELL + i);
	Logger.log("sheet name:" + sheetName + ", gid:" + sheetGid);
	if (!String(sheetGid).length) {
	    break;      
	}
	var docId = SpreadsheetApp.getActiveSpreadsheet().getId();
	var url = SPREADSHEET_URL + docId + "&exportFormat=" + EXPORT_FORMAT + "&gid=" + sheetGid;
	var auth = "AuthSub token=\"" + getConfig(row, TOKEN_CELL) + "\"";
	var res = UrlFetchApp.fetch(url, {headers: {Authorization: auth}});
	var content = res.getContent();
	var responseCode = res.getResponseCode();
	if (responseCode != 200 || res.getContentText().indexOf("/ServiceLoginAuth") != -1) {
	    Logger.log("Fetch url:" + url + " failed with " + responseCode);
	    Browser.msgBox("Error occurred when exporting spreadsheet to pdf, it might be caused by auth token being expired");
	    return;
	}
	attachments.push({fileName:sheetName +"_" + i + "." + EXPORT_FORMAT, content: content});
    }
    var bcc = getConfig(row, BCC_CELL);
    Logger.log("BCC to:" + bcc);
    MailApp.sendEmail(getConfig(row, EMAIL_CELL), getConfig(row, SUBJECT_CELL),
		      getConfig(row, BODY_CELL), {attachments:attachments, bcc: bcc});
}

function getConfig(row, col) {
    var configSheet = getConfigSheet();
    return configSheet.getRange(row, col).getValue();
}

function getConfigSheet() {
    var name = "script_config";
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
    if (!sheet) {
	sheet = SpreadsheetApp.getActiveSpreadsheet().insertSheet(name);
	Logger.log("Created sheet " + name);
    }
    return sheet;
}

​
