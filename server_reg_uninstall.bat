@ECHO OFF
CD /d %~dp0

SET SERVER_PROG="%~dp0CheckRerun.vbs"
schtasks /delete /tn CheckRerun /F