function onInstall(e) {
	onOpen(e)
}

function onOpen(e) {
	var addonMenu = SpreadsheetApp.getUi().createAddonMenu()

	addonMenu.addItem('Transcriber Management', 'onShowTxManagementSidebar')
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
	var sheet = getSheet('Transcribers')
	var headerCol = {}

	if (sheetName === 'Transcribers') {
		var headers = ['Name', 'Email', 'SheetID']
	}

	var _range = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()
	_.each(headers, function (key) {
		headerCol[key] = _.findIndex(_range[0], function (value) { return key === value }) + 1
	})
	return headerCol
}

function newTranscriber(newTx) {
	var sheet = getSheet('Transcribers')
	var headerIndex = getHeaderIndex('Transcribers')
	var newRowNum = sheet.getLastRow() + 1

	sheet.getRange(newRowNum, headerIndex['Name'], 1, 1).setValue(newTx.name)
	sheet.getRange(newRowNum, headerIndex['Email'], 1, 1).setValue(newTx.email)

}
