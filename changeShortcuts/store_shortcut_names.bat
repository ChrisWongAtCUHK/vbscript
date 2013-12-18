@ECHO OFF
@REM Constants
SET SHORTCUTS="D:\GoogleChromePortableMS\chrome-shortcuts-app\*.lnk"
SET TMPTXT=%USERPROFILE%\Documents\tmp.txt

@REM Create necessary file
COPY NUL %TMPTXT%

SETLOCAL ENABLEDELAYEDEXPANSION
SET /A INDEX=1
@REM Loop through the target directory recursively 
FOR /F "tokens=* delims=" %%A IN ('DIR /S /B %SHORTCUTS%') DO (
	ECHO %%A >> %TMPTXT%
	SET FILEPATH=
	SET FILENAME=
	FOR %%P IN ("%%A") DO SET FILEPATH=%%~dpP
	FOR %%N IN ("%%A") DO SET FILENAME=%%~nxN
	MOVE "!FILEPATH!!FILENAME!" "!FILEPATH!!INDEX!.lnk"
	ECHO !INDEX!
	SET /A INDEX+=1
)

ECHO INDEX=%INDEX%
ENDLOCAL

@REM	ECHO !INDEX!:%%A >> %TMPTXT%
@REM		1.	Output to tmp text file
@REM	SET FILEPATH=
@REM	FOR %%I IN ("%%A") DO ECHO %%~dpI
@REM		1.	Truncate the directory name without filename
@REM		2.	IMPORTANT: "", in order to include space
@REM	SET FILENAME=
@REM		FOR %%N IN ("%%A") DO SET FILENAME=%%~nxN
@REM		1.	Truncate the file name without directory
@REM		2.	IMPORTANT: "", in order to include space
@REM	MOVE "!FILEPATH!!FILENAME!" "!FILEPATH!!INDEX!.lnk"
@REM		1.	Change file
@REM	SET /A INDEX+=1
@REM		1.	Counter

@REM http://stackoverflow.com/questions/659647/how-to-get-folder-path-from-file-path-with-cmd
