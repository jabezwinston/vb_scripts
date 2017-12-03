' Author : Jabez Winston
'  
' Date : 3-12-2017

wordFilename = "winston.docx"
fileFullPath = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\" & wordFilename

Set objWord = CreateObject("Word.Application")
objWord.Visible = True

Set objDoc = objWord.Documents.Add()
Set objSelection = objWord.Selection

objSelection.Font.Name = "Arial"
objSelection.Font.Size = "18"
objSelection.TypeText "Hello World !!"
wscript.sleep 1000
objSelection.TypeParagraph()

objSelection.Font.Size = "14"
objSelection.TypeText "" & Now()
wscript.sleep 1000

objSelection.Font.Size = "10"
wscript.sleep 2000

objDoc.SaveAs(fileFullPath)
objWord.Quit