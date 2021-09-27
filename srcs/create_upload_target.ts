function manual_create_upload_target() {
	let confirmation = Browser.msgBox('ドメインリスト抽出処理', '本当に実行しますか？', Browser.Buttons.OK_CANCEL);
	if (confirmation == "cancel") {
		return ;
	}
	create_upload_target();
}

function create_upload_target() {
	const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
	const SHEET = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('UploadInfo');
	SHEET.clear();
	if (SHEET.getFilter())
		SHEET.getFilter().remove();
	SHEET.getRange('A1').setValue('サーバー番号').setBackground('#c9daf8');
	SHEET.getRange('B1').setValue('ドメイン名').setBackground('#c9daf8');
	SHEET.getRange('C1').setValue('VGSEO納品データ').setBackground('#c9daf8');
	SHEET.getRange('D1').setValue('データ番号').setBackground('#c9daf8');
	SHEET.getRange('E1').setValue('アップロード数').setBackground('#f9cb9c');
	SHEET.getRange('F1').setValue('=ArrayFormula(match(0,len(B2:B),0))-1');
	SHEET.getRange('H1').setValue(Utilities.formatDate(new Date(), 'JST', 'yyyy-MM-dd'))
		.setBackground('#efefef');
	SHEET.getRange('A1:F1').setFontFamily('Meiryo')
		.setFontWeight('bold')
	SHEET.getRange('A1:H1').setFontFamily('Meiryo')
		.setHorizontalAlignment('center')
		.setVerticalAlignment('middle');
	for (let col = 1; col <= 8; col++) {
		if (col == 1)
			SHEET.setColumnWidth(col, 110);
		else if (col == 2 || col == 3)
			SHEET.setColumnWidth(col, 200);
		else if (col == 5)
			SHEET.setColumnWidth(col, 120);
		else if (col == 6)
			SHEET.setColumnWidth(col, 70);
		else
			SHEET.setColumnWidth(col, 100);
	}
	SHEET.setRowHeight(1, 40);
	SHEET.setFrozenRows(1);
	let targetDomainList: Array<Array<string>>;
	targetDomainList = get_upload_target()
	if (targetDomainList.length < 1)
		return ;
	SHEET.getRange(2, 1, targetDomainList.length, 4)
		.setValues(targetDomainList)
		.setFontFamily('Meiryo');
	SHEET.getRange(1, 1, SHEET.getLastRow(), 4).createFilter();
}

function get_upload_target() {
	const SPREADSHEET_ID = PropertiesService.getScriptProperties().getProperty('SERVER123_SSID');
	const TARGET_SHEET = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('Main');
	
	const LAST_ROW = TARGET_SHEET.getLastRow();
	let domainList: Array<Array<string>> = TARGET_SHEET.getRange(`A3:Q${LAST_ROW}`).getValues();
	domainList = domainList.filter(data => !data[11])
		.map(data => [data[0], data[5], data[15], data[16]]);
	return domainList;
}
