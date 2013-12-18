@ECHO OFF
@REM Constants
SET TMPTXT=%USERPROFILE%\Documents\tmp.txt

SETLOCAL ENABLEDELAYEDEXPANSION
SET /A INDEX=1
@REM Loop through the target directory recursively 
FOR /F "tokens=* delims=" %%A IN ('TYPE %TMPTXT%') DO (
	SET FILEPATH=
	SET FILENAME=
	FOR %%P IN ("%%A") DO SET FILEPATH=%%~dpP
	FOR %%N IN ("%%A") DO SET FILENAME=%%~nN
	ECHO "!FILEPATH!!INDEX!.lnk"
	ECHO "!FILEPATH!!FILENAME!.lnk"
	MOVE "!FILEPATH!!INDEX!.lnk" "!FILEPATH!!FILENAME!.lnk"
	SET /A INDEX+=1
)
ECHO INDEX=%INDEX%
ENDLOCAL


@REM http://stackoverflow.com/questions/659647/how-to-get-folder-path-from-file-path-with-cmd