function onInstall(e) {
	onOpen(e)
}

function onOpen(e) {
	var addonMenu = SpreadsheetApp.getUi().createAddonMenu()

	addonMenu.addItem('Tasks Management', 'onShowTasksManagementSidebar')
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

function onShowTasksManagementSidebar() {
	var html = HtmlService.createTemplateFromFile('tasksManagement')
	html.mode = 'addon'
	SpreadsheetApp.getUi()
		.showSidebar(html.evaluate()
			.setSandboxMode(HtmlService.SandboxMode.IFRAME)
			.setTitle('Tasks Management'))
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
		headers = ['File', 'MP3 Source', 'Original Source', 'Audio Length', 'Speakers', 'Transcriber 1', 'Transcriber 1 Status', 'Transcriber 1 Distribution Date', 'Transcriber 1 Result', 'Transcriber 1 Time Taken', 'Transcriber 1 Tool', 'Transcriber 1 Issues/Bugs', 'Transcriber 1 Notes', 'Transcriber 2', 'Transcriber 2 Status','Transcriber 2 Distribution Date', 'Transcriber 2 Result', 'Transcriber 2 Time Taken', 'Transcriber 2 Tool', 'Transcriber 2 Issues/Bugs','Transcriber 2 Notes']
	}

	var _range = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()
	_.each(headers, function (key) {
		headerCol[key] = _.findIndex(_range[0], function (value) { return key === value }) + 1
	})
	return headerCol
}

function getHeaderIndexTx(fileId) {
	var _ = LodashGS.load()
	var sheet = SpreadsheetApp.openById(fileId).getSheetByName('Tasks')
	var headerCol = {}
	var headers = ['File','MP3 Source',	'Original Source','Speakers', 'Result',	'Time Taken',	'Tool',	'Issues/Bugs']

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

function distributeTasks() {
	var _ = LodashGS.load()
	var tasksMasterSheet = getSheet('TasksMaster')
	var headerIndexTasksMaster = getHeaderIndex('TasksMaster')

	//Distribute Tasks that has a name but does not have distribution dates
	var dataFile = tasksMasterSheet.getRange(2, headerIndexTasksMaster['File'], tasksMasterSheet.getLastRow() - 1, 1).getValues()
	var dataMP3 = tasksMasterSheet.getRange(2, headerIndexTasksMaster['MP3 Source'], tasksMasterSheet.getLastRow() - 1, 1).getValues()
	var dataOriginal = tasksMasterSheet.getRange(2, headerIndexTasksMaster['Original Source'], tasksMasterSheet.getLastRow() - 1, 1).getValues()
	var dataSpeakers = tasksMasterSheet.getRange(2, headerIndexTasksMaster['Speakers'], tasksMasterSheet.getLastRow() - 1, 1).getValues()
	var dataAudioLength = tasksMasterSheet.getRange(2, headerIndexTasksMaster['Audio Length'], tasksMasterSheet.getLastRow() - 1, 1).getValues()
	var dataToolTx1 = tasksMasterSheet.getRange(2, headerIndexTasksMaster['Transcriber 1 Tool'], tasksMasterSheet.getLastRow() - 1, 1).getValues()
	var dataToolTx2 = tasksMasterSheet.getRange(2, headerIndexTasksMaster['Transcriber 2 Tool'], tasksMasterSheet.getLastRow() - 1, 1).getValues()

	var tx1 = tasksMasterSheet.getRange(2, headerIndexTasksMaster['Transcriber 1'], tasksMasterSheet.getLastRow() - 1, 1).getValues()
	var distributionDatesTx1 = tasksMasterSheet.getRange(2, headerIndexTasksMaster['Transcriber 1 Distribution Date'], tasksMasterSheet.getLastRow() - 1, 1).getValues()
	var tx1Combined = tx1.map(function (txRow, index) { return {rowArrayIndex: index, name: txRow[0], distDate: distributionDatesTx1[index][0]} })
	var tx2 = tasksMasterSheet.getRange(2, headerIndexTasksMaster['Transcriber 2'], tasksMasterSheet.getLastRow() - 1, 1).getValues()
	var distributionDatesTx2 = tasksMasterSheet.getRange(2, headerIndexTasksMaster['Transcriber 2 Distribution Date'], tasksMasterSheet.getLastRow() - 1, 1).getValues()
	var tx2Combined = tx2.map(function (txRow, index) { return {rowArrayIndex: index, name: txRow[0], distDate: distributionDatesTx2[index][0]} })

	var requiredDistributionTx1 = _.filter(tx1Combined, function (obj) { return (obj.name !== '' && obj.distDate === '')})
	var requiredDistributionTx2 = _.filter(tx2Combined, function (obj) { return (obj.name !== '' && obj.distDate === '')})

	var sheet = getSheet('Transcribers')
	var headerIndexTx = getHeaderIndex('Transcribers')
	var txNames = sheet.getRange(2, headerIndexTx['Name'], sheet.getLastRow() - 1, 1).getValues()
	var txFileId = sheet.getRange(2, headerIndexTx['FileID'], sheet.getLastRow() -1 , 1).getValues()
	var txLibrary = txNames.map(function (nameRow, index) { return {name: nameRow[0], fileId: txFileId[index][0]} })

	for (var i = requiredDistributionTx1.length - 1; i >= 0; i--) {
		var thisTxName = requiredDistributionTx1[i].name
		var thisTxFileId = _.find(txLibrary, function (o) { return o.name === thisTxName }).fileId
		var thisTxSheet = SpreadsheetApp.openById(thisTxFileId).getSheetByName('Tasks')
		var headerIndexTxSheet = getHeaderIndexTx(thisTxFileId)
		var newRowIndex = thisTxSheet.getLastRow() + 1

		thisTxSheet.getRange(newRowIndex, headerIndexTxSheet['File'], 1, 1).setValue(dataFile[requiredDistributionTx1[i].rowArrayIndex])
		thisTxSheet.getRange(newRowIndex, headerIndexTxSheet['MP3 Source'], 1, 1).setValue(dataMP3[requiredDistributionTx1[i].rowArrayIndex])
		thisTxSheet.getRange(newRowIndex, headerIndexTxSheet['Original Source'], 1, 1).setValue(dataOriginal[requiredDistributionTx1[i].rowArrayIndex])
		thisTxSheet.getRange(newRowIndex, headerIndexTxSheet['Speakers'], 1, 1).setValue(dataSpeakers[requiredDistributionTx1[i].rowArrayIndex])
		thisTxSheet.getRange(newRowIndex, headerIndexTxSheet['Tool'], 1, 1).setValue(dataToolTx1[requiredDistributionTx1[i].rowArrayIndex])

		//Use ImportRange formula to get updates from transcriber's tasks list
		tasksMasterSheet.getRange(requiredDistributionTx1[i].rowArrayIndex + 2, headerIndexTasksMaster['Transcriber 1 Result'], 1, 1)
			.setFormula('=ImportRange("https://docs.google.com/spreadsheets/d/' + thisTxFileId + '", "Tasks!" & address(' + newRowIndex + ',' + headerIndexTxSheet['Result']+ '))')
		tasksMasterSheet.getRange(requiredDistributionTx1[i].rowArrayIndex + 2, headerIndexTasksMaster['Transcriber 1 Time Taken'], 1, 1)
			.setFormula('=ImportRange("https://docs.google.com/spreadsheets/d/' + thisTxFileId + '", "Tasks!" & address(' + newRowIndex + ',' + headerIndexTxSheet['Time Taken']+ '))')
		tasksMasterSheet.getRange(requiredDistributionTx1[i].rowArrayIndex + 2, headerIndexTasksMaster['Transcriber 1 Issues/Bugs'], 1, 1)
			.setFormula('=ImportRange("https://docs.google.com/spreadsheets/d/' + thisTxFileId + '", "Tasks!" & address(' + newRowIndex + ',' + headerIndexTxSheet['Issues/Bugs']+ '))')
		tasksMasterSheet.getRange(requiredDistributionTx1[i].rowArrayIndex + 2, headerIndexTasksMaster['Transcriber 1 Notes'], 1, 1)
			.setFormula('=ImportRange("https://docs.google.com/spreadsheets/d/' + thisTxFileId + '", "Tasks!" & address(' + newRowIndex + ',' + headerIndexTxSheet['Notes']+ '))')

		//Set Distribution Date
		tasksMasterSheet.getRange(requiredDistributionTx1[i].rowArrayIndex + 2, headerIndexTasksMaster['Transcriber 1 Distribution Date'], 1, 1).setValue(new Date())
		tasksMasterSheet.getRange(requiredDistributionTx1[i].rowArrayIndex + 2, headerIndexTasksMaster['Transcriber 1 Status'], 1, 1).setValue('In Progress')

	}

	for (var i = requiredDistributionTx2.length - 1; i >= 0 ; i--) {
		var thisTxName = requiredDistributionTx2[i].name
		var thisTxFileId = _.find(txLibrary, function (o) { return o.name === thisTxName }).fileId
		var thisTxSheet = SpreadsheetApp.openById(thisTxFileId)
		var headerIndexTxSheet = getHeaderIndexTx(thisTxFileId)
		var newRowIndex = thisTxSheet.getLastRow() + 1

		thisTxSheet.getRange(newRowIndex, headerIndexTxSheet['File'], 1, 1).setValue(dataFile[requiredDistributionTx2[i].rowArrayIndex])
		thisTxSheet.getRange(newRowIndex, headerIndexTxSheet['MP3 Source'], 1, 1).setValue(dataMP3[requiredDistributionTx2[i].rowArrayIndex])
		thisTxSheet.getRange(newRowIndex, headerIndexTxSheet['Original Source'], 1, 1).setValue(dataOriginal[requiredDistributionTx2[i].rowArrayIndex])
		thisTxSheet.getRange(newRowIndex, headerIndexTxSheet['Speakers'], 1, 1).setValue(dataSpeakers[requiredDistributionTx2[i].rowArrayIndex])
		thisTxSheet.getRange(newRowIndex, headerIndexTxSheet['Tool'], 1, 1).setValue(dataToolTx2[requiredDistributionTx2[i].rowArrayIndex])

		//Use ImportRange formula to get updates from transcriber's tasks list
		tasksMasterSheet.getRange(requiredDistributionTx2[i].rowArrayIndex + 2, headerIndexTasksMaster['Transcriber 2 Result'], 1, 1)
			.setFormula('=ImportRange("https://docs.google.com/spreadsheets/d/' + thisTxFileId + '", "Tasks!" & address(' + newRowIndex + ',' + headerIndexTxSheet['Result']+ '))')
		tasksMasterSheet.getRange(requiredDistributionTx2[i].rowArrayIndex + 2, headerIndexTasksMaster['Transcriber 2 Time Taken'], 1, 1)
			.setFormula('=ImportRange("https://docs.google.com/spreadsheets/d/' + thisTxFileId + '", "Tasks!" & address(' + newRowIndex + ',' + headerIndexTxSheet['Time Taken']+ '))')
		tasksMasterSheet.getRange(requiredDistributionTx2[i].rowArrayIndex + 2, headerIndexTasksMaster['Transcriber 2 Issues/Bugs'], 1, 1)
			.setFormula('=ImportRange("https://docs.google.com/spreadsheets/d/' + thisTxFileId + '", "Tasks!" & address(' + newRowIndex + ',' + headerIndexTxSheet['Issues/Bugs']+ '))')
		tasksMasterSheet.getRange(requiredDistributionTx2[i].rowArrayIndex + 2, headerIndexTasksMaster['Transcriber 2 Notes'], 1, 1)
			.setFormula('=ImportRange("https://docs.google.com/spreadsheets/d/' + thisTxFileId + '", "Tasks!" & address(' + newRowIndex + ',' + headerIndexTxSheet['Notes']+ '))')

		//Set Distribution Date
		tasksMasterSheet.getRange(requiredDistributionTx2[i].rowArrayIndex + 2, headerIndexTasksMaster['Transcriber 2 Distribution Date'], 1, 1).setValue(new Date())
		tasksMasterSheet.getRange(requiredDistributionTx2[i].rowArrayIndex + 2, headerIndexTasksMaster['Transcriber 2 Status'], 1, 1).setValue('In Progress')
	}

	SpreadsheetApp.getActiveSpreadsheet().toast('Tasks Assignment Completed', '', 3)
}

function unassignTask(fileName, whichTx) {

	//Find the respective file from the Transcriber's and delete the row

	//Update the row values in Master
}

function saveSettings(newSettings) {
	PropertiesService.getUserProperties().setProperties(newSettings)
	SpreadsheetApp.getActiveSpreadsheet().toast('Settings have been saved', '', 3)
	return null
}

function getSavedProperties() {
	return PropertiesService.getUserProperties().getProperties()
}
