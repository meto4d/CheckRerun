@ECHO OFF
CD /d %~dp0

rem é¿çsÇ∑ÇÈä‘äu(1Å`999ï™)
SET /A TIME_INTERVAL=1

SET SERVER_PROG=\"wscript"\"%~dp0CheckRerun.vbs"
::echo %SERVER_PROG%
schtasks /create /tn CheckRerun /tr %SERVER_PROG% /sc MINUTE /mo %TIME_INTERVAL%