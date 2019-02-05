' Set the autorun parameters.
With CreateObject("WScript.Shell").Environment("PROCESS")
	.Item("APP_IS_AUTORUN_MODE") = "TRUE"
	.Item("APP_INPUT_TABLE_FILE_PATH") = "H:\INPUT_TABLE.xlsx"
	.Item("APP_BODY_TEMPLATE_FILE_PATH") = "H:\BODY_TEMPLATE.txt"
End With

' Run the main project workbook.
CreateObject("Excel.Application") _
	.Workbooks.Open(CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\MailSender.xlsm")
