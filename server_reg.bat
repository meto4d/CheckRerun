@ECHO OFF
CD /d %~dp0

rem ���s����Ԋu(1�`999��)
SET /A TIME_INTERVAL=5

SET SERVER_PROG=\"wscript"\"%~dp0CheckRerun.vbs"
::echo %SERVER_PROG%
schtasks /create /tn CheckRerun /tr %SERVER_PROG% /sc MINUTE /mo %TIME_INTERVAL%