@ECHO OFF
CD /d %~dp0

rem 実行する間隔(1〜999分)
SET /A TIME_INTERVAL=5

SET SERVER_PROG=\"wscript"\"%~dp0CheckRerun.vbs"
::echo %SERVER_PROG%
schtasks /create /tn CheckRerun /tr %SERVER_PROG% /sc MINUTE /mo %TIME_INTERVAL%