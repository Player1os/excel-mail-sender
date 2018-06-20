@ECHO OFF

:: Set the project password.
SET APP_DEBUG_PASSWORD=tele$ExcelMailSender

:: Run the main project workbook.
CALL "%~dp0MailSender.xlsm"
