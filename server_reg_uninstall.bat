@ECHO OFF
CD /d %~dp0

rem ���s����Ԋu(��)
SET /A TIME_INTERVAL=5

SET SERVER_PROG="%~dp0CheckRerun.vbs"
schtasks /delete /tn CheckRerun /F