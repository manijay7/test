REM source: http://stackoverflow.com/a/11252731/715608

:: Iterate over all xlx* files and call script
:: FOR /f "delims=" %%i IN ('DIR *.xls* /b') DO to-csv.vbs "%%i" "%%~ni.csv"

:: Set currentDir to dir with script
SET currentDir=%~dp0to_csv.vbs

cscript %currentDir% "%1" "%2.csv"
