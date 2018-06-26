@ECHO OFF

:: Switch to the deploy directory.
PUSHD \\kedata\Data1\B2B_Business_Inteligence\Osama Hassanein\ExcelMailSender

:: Set the autorun parameters.
SET APP_IS_AUTORUN_MODE=TRUE
SET APP_INPUT_TABLE_FILE_PATH=H:\INPUT_TABLE.xlsx
SET APP_BODY_TEMPLATE_FILE_PATH=H:\BODY_TEMPLATE.txt

:: Run the main project workbook.
CALL "MailSender.xlsm"

:: Return to the original working directory.
POPD
