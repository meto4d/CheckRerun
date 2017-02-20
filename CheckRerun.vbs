'//-----------------------------
'// �w�肵���v���O���������s����Ă��Ȃ���΃v���O�����̍ċN�����s���A
'// �����Ȃ��ł���΋����I����A�v���O�����̍ċN�����s���v���O�����̍ċN�����s���v���O����
'// 2017/02/18 ver 0.91
'// @meto4d : Ruth
'//-----------------------------
Option Explicit


' �����ɊĎ����s��server�v���O������"�v���Z�X��"����͂���
Const SERVER_PROG = "cmd.exe"

' �����Ɏ��s���s��server�v���O������"�ꏊ"����͂���
Const SERVER_DIR = "C:\Windows\System32"

' �����Ɏ��s���s��server�v���O������"���O"����͂���
Const SERVER_EXEC = "notepad.exe"

' �N�����Ă��Ȃ��������ɁA�ċN�������������o�͂��Ă����t�@�C��
Const LOG_PROG = "process_test.txt"



'//~~~~~~~~�����J�n~~~~~~~~

Dim objWShell
Set objWShell = CreateObject("WScript.Shell")
Dim ThisDir : ThisDir = Replace(WScript.ScriptFullName,WScript.ScriptName,"")
Dim RetExiProData
RetExiProData = ProcessExists(SERVER_PROG)

' �v���Z�X�����s����Ă����
If RetExiProData(0) Then
	
	' �e�v���Z�X���ŉ����Ȃ����ǂ�����������x�₢���킹����ŏ����������I��
	Dim RetProId, oExec
	For Each RetProId In RetExiProData(1)
		
		If RetProId <> 0 Then
			
			' ���O�o�͗p
			Call SysWriter(""""& SERVER_PROG & """ : " & RetProId & " : �����Ȃ�", ThisDir & LOG_PROG)
			
			'�v���O���������I��
			Set oExec = objWShell.Exec("WScript.exe " & ThisDir & "HideExec.vbs taskkill.exe /f /t /fi ""status eq not responding"" /pid " & RetProId)
			oExec.StdIn.Close
			
			WScript.Sleep 100
			
			'�v���O�����Ď��s
			objWShell.CurrentDirectory = SERVER_DIR
			objWShell.run SERVER_EXEC, 7
			objWShell.CurrentDirectory = ThisDir
			
		End If
		
	Next
	
' �v���Z�X�����s����Ă��Ȃ����
Else
	
	' ���O�o�͗p
	Call SysWriter(""""& SERVER_PROG & """ : �N��NG", ThisDir & LOG_PROG)
	
	' SERVER_DIR�����s:�A�N�e�B�u��؂�ւ������Ŏ��s������(7):Debug����1
	objWShell.CurrentDirectory = SERVER_DIR
	objWShell.run SERVER_EXEC, 7
	objWShell.CurrentDirectory = ThisDir
	
End If

'//~~~~~~~~�����I���~~~~~~~~


'//-------------
'// ���O�o�͗p
'// Call SysWriter("���O���b�Z�[�W", "test_log.txt")
'//-------------
Sub SysWriter(str,path)
 
	Dim fso, fi
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	'�t�@�C�����J��
	 '���������݂��Ȃ��ꍇ�ɂ͍쐬����
	Set fi = fso.OpenTextFile(path, 8, true)
	
	If Err.Number = 0 Then
		' ���O�o��
		fi.WriteLine(Date() & " " & Time() & " - " & str)
		fi.Close
	Else
		' ���O�o�͗p�t�@�C���̃I�[�v���Ɏ��s�����ꍇ
		WScript.Echo "�t�@�C���I�[�v���G���[: " & Err.Description
	End If
	
	Set fi = Nothing
 
End Sub


'//-----------------------------
'// �v���Z�X���N�����Ă��邩���ׂ�(�����true)
'//-----------------------------
Function ProcessExists(ProcessName)
	'�Ԃ�l�p Array("�v���O�����̗L��BOOL", Array( Process ID,,, (�������Ȃ����0) )) )
	Dim RetData(1)
	
	'�v���Z�X�`�F�b�N
	Dim Service, QfeSet, Qfe, QproId
	
	Set Service = WScript.CreateObject("WbemScripting.SWbemLocator").ConnectServer

    Set QfeSet = Service.ExecQuery("Select * From Win32_Process Where Caption='" & ProcessName & "'")
	
	QproId = 0
	For Each Qfe In QfeSet
        QproId = Qfe.ProcessId
        Exit For
    Next
	
	RetData(0) = QproId <> 0
	
	If RetData(0) Then
	
		Dim ProcArray
		Set ProcArray = CreateObject("System.Collections.ArrayList")
		
	'�����Ȃ��̃v���O������T������
		
		'�^�X�N�����ׂăI�u�W�F�N�g�ɓ����
		Dim objExec
		Set objExec = objWShell.Exec("WScript.exe " & ThisDir & "HideExec.vbs tasklist /fi ""status eq not responding""")
		objExec.StdIn.Close
		
		'�o�͌���
		Dim objStdOut
		Set objStdOut = objExec.StdOut
		
		'�o�͌��ʂ���T��
		Dim strLine, strTest
		Do Until objStdOut.AtEndOfStream
		
			strLine = objStdOut.ReadLine
			If (InStr(strLine, SERVER_PROG) <> 0) Then
			
				'�Y���v���O���������݂����ꍇ
				strTest = strLine
				Call SysWriter(strTest, "test_log.txt")
				strTest = Right(strTest, Len(strTest) - Len(SERVER_PROG))
				ProcArray.Add Split(LTrim(strTest), " ")(0)
				
			End If
			
		Loop
		
		objStdOut.Close
		
		RetData(1) = ProcArray.ToArray()
	
	Else
		
		RetData(1) = Array("0")
	
	End If
	
	ProcessExists = RetData

End Function
