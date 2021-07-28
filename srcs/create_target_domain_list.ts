function create_target_domain_list() {
	let confirmation = Browser.msgBox('ドメインリスト抽出処理', '本当に実行しますか？', Browser.Buttons.OK_CANCEL);
	if (confirmation == "cancel") {
		return;
	}

	const SHEET = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
	SHEET.clear();
	if (SHEET.getFilter())
		SHEET.getFilter().remove();
	SHEET.getRange('A1').setValue('サーバー番号').setBackground('#c9daf8');
	SHEET.getRange('B1').setValue('ドメイン名').setBackground('#c9daf8');
	SHEET.getRange('C1').setValue('Size').setBackground('#c9daf8');
	SHEET.getRange('D1').setValue('=ArrayFormula(match(0,len(B2:B),0))-1');
	SHEET.getRange('A1:N1').setFontFamily('Meiryo')
		.setFontWeight('bold')
		.setHorizontalAlignment('center')
		.setVerticalAlignment('middle');
	for (let col = 1; col <= 4; col++) {
		if (col == 1)
			SHEET.setColumnWidth(col, 110);
		else if (col == 2)
			SHEET.setColumnWidth(col, 200);
		else
			SHEET.setColumnWidth(col, 70);
	}
	SHEET.setRowHeight(1, 40);
	SHEET.setFrozenRows(1);
	let targetDomainList: Array<Array<string>>;
	if (SHEET.getSheetName() == 'RemoveInfoFtp')
		targetDomainList = get_target_domain_list('登録中ドメイン（FTPサーバー）')
	else if (SHEET.getSheetName() == 'RemoveInfo123')
		targetDomainList = get_target_domain_list('登録中ドメイン（123サーバー）')
	SHEET.getRange(2, 1, targetDomainList.length, 2)
		.setValues(targetDomainList)
		.setFontFamily('Meiryo');
	SHEET.getRange(1, 1, SHEET.getLastRow(), 2).createFilter();
}

function get_target_domain_list(sheetName: string) {
	const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('UNDER_CONTRACT_SSID');
	const TARGET_SHEET = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(sheetName);
	
	const LIST_SIZE: number = TARGET_SHEET.getRange('E1').getValue();
	let domainList: Array<Array<string>> = TARGET_SHEET.getRange(2, 1, LIST_SIZE, 3).getValues();
	domainList = domainList.filter(data => data[2] == false)
		.map(data => [data[0], data[1]]);
	return domainList;
}
