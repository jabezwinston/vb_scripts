'https://gist.github.com/simply-coded/758e2557ccd1a55b46765d8bb1099ec6

excelFilename = "winston.xlsx"
fileFullPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\" & excelFilename

'create the excel object
	Set objExcel = CreateObject("Excel.Application") 

'view the excel program and file, set to false to hide the whole process
	objExcel.Visible = True 

'add a new workbook
	Set objWorkbook = objExcel.Workbooks.Add 

'set a cell value at row 3 column 5
	objExcel.Cells(1,1).Value = "Jabez"
	wscript.sleep 2000

'change a cell value
	objExcel.Cells(2,1).Value = "Winston"
	wscript.sleep 2000
	
'delete a cell value
	objExcel.Cells(3,5).Value = ""

'get a cell value and set it to a variable
	 a = objExcel.Cells(1,1).Value
	 wscript.echo a
	 wscript.sleep 2000

'save the new excel file (make sure to change the location) 'xls for 2003 or earlier
	objWorkbook.SaveAs(fileFullPath) 

'close the workbook
	objWorkbook.Close 

'exit the excel program
	objExcel.Quit

'release objects
	Set objExcel = Nothing
	Set objWorkbook = Nothing