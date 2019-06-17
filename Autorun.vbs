Option Explicit

' Set the autorun parameters.
With CreateObject("WScript.Shell").Environment("PROCESS")
	.Item("APP_IS_AUTORUN_MODE") = "TRUE"
	.Item("APP_INPUT_TABLE_FILE_PATH") = "H:\INPUT_TABLE.xlsx"
	.Item("APP_BODY_TEMPLATE_FILE_PATH") = "H:\BODY_TEMPLATE.txt"
	.Item("APP_SENDER_ACCOUNT_INDEX") = "1"
End With

' Run the main project workbook.
Call CreateObject("Excel.Application").Workbooks.Open( _
	CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName) & "\MailSender.xlsm" _
)
