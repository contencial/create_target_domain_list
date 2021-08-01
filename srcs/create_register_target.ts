function manual_create_register_target() {
	let confirmation = Browser.msgBox('ドメインリスト抽出処理', '本当に実行しますか？', Browser.Buttons.OK_CANCEL);
	if (confirmation == "cancel") {
		return ;
	}
	create_register_target();
}

function create_register_target() {
	const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
	const SHEET = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('RegisterInfo');
	SHEET.clear();
	if (SHEET.getFilter())
		SHEET.getFilter().remove();
	SHEET.getRange('A1').setValue('サーバー番号').setBackground('#c9daf8');
	SHEET.getRange('B1').setValue('ドメイン名').setBackground('#c9daf8');
	SHEET.getRange('C1').setValue('Size').setBackground('#c9daf8');
	SHEET.getRange('D1').setValue('=ArrayFormula(match(0,len(B2:B),0))-1');
	SHEET.getRange('F1').setValue(Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd'))
		.setBackground('#efefef');
	SHEET.getRange('A1:D1').setFontFamily('Meiryo')
		.setFontWeight('bold')
	SHEET.getRange('A1:F1').setFontFamily('Meiryo')
		.setHorizontalAlignment('center')
		.setVerticalAlignment('middle');
	for (let col = 1; col <= 6; col++) {
		if (col == 1)
			SHEET.setColumnWidth(col, 110);
		else if (col == 2)
			SHEET.setColumnWidth(col, 200);
		else if (col == 3 || col == 4)
			SHEET.setColumnWidth(col, 70);
		else
			SHEET.setColumnWidth(col, 100);
	}
	SHEET.setRowHeight(1, 40);
	SHEET.setFrozenRows(1);
	let targetDomainList: Array<Array<string>>;
	targetDomainList = get_register_target()
	if (targetDomainList.length < 1)
		return ;
	SHEET.getRange(2, 1, targetDomainList.length, 2)
		.setValues(targetDomainList)
		.setFontFamily('Meiryo');
	SHEET.getRange(1, 1, SHEET.getLastRow(), 2).createFilter();
}

function get_register_target() {
	const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SERVER123_SSID');
	const TARGET_SHEET = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Main');
	
	const LAST_ROW = TARGET_SHEET.getLastRow();
	let domainList: Array<Array<string>> = TARGET_SHEET.getRange(`A3:M${LAST_ROW}`).getValues();
	domainList = domainList.filter(data => !data[12])
		.map(data => [data[0], data[4]]);
	return domainList;
}
