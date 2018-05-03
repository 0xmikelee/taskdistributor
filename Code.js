function onInstall(e) {
	onOpen(e)
}

function onOpen(e) {
	var addonMenu = SpreadsheetApp.getUi().createAddonMenu()

	addonMenu.addItem('Transcriber Management', 'onShowTxManagementSidebar')
	addonMenu.addItem('Settings', 'onShowSettingsSidebar')
	addonMenu.addToUi()
}

function include(File) {
  return HtmlService.createHtmlOutputFromFile(File).getContent();
};

function onShowTxManagementSidebar() {
	var html = HtmlService.createTemplateFromFile('txManagement')
	html.mode = 'addon'
	SpreadsheetApp.getUi()
		.showSidebar(html.evaluate()
			.setSandboxMode(HtmlService.SandboxMode.IFRAME)
			.setTitle('Transcribers Management'))
}

function onShowSettingsSidebar() {
	var html = HtmlService.createTemplateFromFile('settings')
	html.mode = 'addon'
	SpreadsheetApp.getUi()
		.showSidebar(html.evaluate()
			.setSandboxMode(HtmlService.SandboxMode.IFRAME)
			.setTitle('Settings'))
}

/**
 * used to expose memebers of a namespace
 * @param {string} namespace name
 * @param {method} method name
 */
function exposeRun(namespace, method, argArray) {
  var func = (namespace ? this[namespace][method] : this[method])
  if (argArray && argArray.length) {
    return func.apply(this, argArray)
  } else {
    return func()
  }
}

function getSheet(sheetName) {
	return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName)
}

function getHeaderIndex(sheetName) {
	var _ = LodashGS.load()
	var sheet = getSheet(sheetName)
	var headerCol = {}
	var headers = []

	if (sheetName === 'Transcribers') {
		headers = ['Name', 'Email', 'FileID']
	} else if (sheetName === 'TasksMaster') {
		headers = ['File', 'MP3 Source', 'Original Source', 'Speakers', 'Transcriber 1', 'Transcriber 1 Distribution Date', 'Transcriber 1 Result', 'Transcriber 1 Time Taken', 'Transcriber 1 Tool', 'Transcriber 1 Issues/Bugs', 'Transcriber 2', 'Transcriber 2 Distribution Date', 'Transcriber 2 Result', 'Transcriber 2 Time Taken', 'Transcriber 2 Tool', 'Transcriber 2 Issues/Bugs']
	}

	var _range = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()
	_.each(headers, function (key) {
		headerCol[key] = _.findIndex(_range[0], function (value) { return key === value }) + 1
	})
	return headerCol
}

function newTranscriber(newTx) {
	var sheet = getSheet('Transcribers')
	var headerIndexTx = getHeaderIndex('Transcribers')
	var newRowNum = sheet.getLastRow() + 1

	//Create a new Spreadsheet File for the new Transcriber
	var templateFileId = PropertiesService.getUserProperties().getProperty('templateFileId')
	var templateFile = DriveApp.getFileById(templateFileId)
	var newFile = templateFile.makeCopy('Transcription Tasks (' + newTx.name + ')')

	sheet.getRange(newRowNum, headerIndexTx['Name'], 1, 1).setValue(newTx.name)
	sheet.getRange(newRowNum, headerIndexTx['Email'], 1, 1).setValue(newTx.email)
	sheet.getRange(newRowNum, headerIndexTx['FileID'], 1, 1).setValue(newFile.getId())

	//Update Transcriber 1, 2 Data Validation Ranges
	var tasksMasterSheet = getSheet('TasksMaster')
	var headerIndexTasksMaster = getHeaderIndex('TasksMaster')
	var rangeTx1 = tasksMasterSheet.getRange(2, headerIndexTasksMaster['Transcriber 1'], tasksMasterSheet.getMaxRows(), 1)
	var rangeTx2 = tasksMasterSheet.getRange(2, headerIndexTasksMaster['Transcriber 2'], tasksMasterSheet.getMaxRows(), 1)
	var rangeTxList = sheet.getRange(2, headerIndexTx['Name'], sheet.getLastRow() - 1, 1)
	var validation = SpreadsheetApp.newDataValidation().requireValueInRange(rangeTxList, true).setAllowInvalid(false).build()
	rangeTx1.setDataValidation(validation)
	rangeTx2.setDataValidation(validation)

	SpreadsheetApp.getActiveSpreadsheet().toast('New Transcriber has been added', '', 3)
	return null
}

function saveSettings(newSettings) {
	PropertiesService.getUserProperties().setProperties(newSettings)
	SpreadsheetApp.getActiveSpreadsheet().toast('Settings have been saved', '', 3)
	return null
}

function getSavedProperties() {
	return PropertiesService.getUserProperties().getProperties()
}
