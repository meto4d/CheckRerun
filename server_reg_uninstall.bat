@ECHO OFF
CD /d %~dp0

rem 実行する間隔(分)
SET /A TIME_INTERVAL=5

SET SERVER_PROG="%~dp0CheckRerun.vbs"
schtasks /delete /tn CheckRerun /F