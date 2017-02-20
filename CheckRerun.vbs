'//-----------------------------
'// 指定したプログラムが実行されていなければプログラムの再起動を行い、
'// 応答なしであれば強制終了後、プログラムの再起動を行うプログラムの再起動を行うプログラム
'// 2017/02/18 ver 0.91
'// @meto4d : Ruth
'//-----------------------------
Option Explicit


' ここに監視を行うserverプログラムの"プロセス名"を入力する
Const SERVER_PROG = "cmd.exe"

' ここに実行を行うserverプログラムの"場所"を入力する
Const SERVER_DIR = "C:\Windows\System32"

' ここに実行を行うserverプログラムの"名前"を入力する
Const SERVER_EXEC = "notepad.exe"

' 起動していなかった時に、再起動した日時を出力してくれるファイル
Const LOG_PROG = "process_test.txt"



'//~~~~~~~~処理開始~~~~~~~~

Dim objWShell
Set objWShell = CreateObject("WScript.Shell")
Dim ThisDir : ThisDir = Replace(WScript.ScriptFullName,WScript.ScriptName,"")
Dim RetExiProData
RetExiProData = ProcessExists(SERVER_PROG)

' プロセスが実行されていれば
If RetExiProData(0) Then
	
	' 各プロセス毎で応答なしかどうかをもう一度問い合わせた上で処理を強制終了
	Dim RetProId, oExec
	For Each RetProId In RetExiProData(1)
		
		If RetProId <> 0 Then
			
			' ログ出力用
			Call SysWriter(""""& SERVER_PROG & """ : " & RetProId & " : 応答なし", ThisDir & LOG_PROG)
			
			'プログラム強制終了
			Set oExec = objWShell.Exec("WScript.exe " & ThisDir & "HideExec.vbs taskkill.exe /f /t /fi ""status eq not responding"" /pid " & RetProId)
			oExec.StdIn.Close
			
			WScript.Sleep 100
			
			'プログラム再実行
			objWShell.CurrentDirectory = SERVER_DIR
			objWShell.run SERVER_EXEC, 7
			objWShell.CurrentDirectory = ThisDir
			
		End If
		
	Next
	
' プロセスが実行されていなければ
Else
	
	' ログ出力用
	Call SysWriter(""""& SERVER_PROG & """ : 起動NG", ThisDir & LOG_PROG)
	
	' SERVER_DIRを実行:アクティブを切り替えず裏で実行させる(7):Debug時は1
	objWShell.CurrentDirectory = SERVER_DIR
	objWShell.run SERVER_EXEC, 7
	objWShell.CurrentDirectory = ThisDir
	
End If

'//~~~~~~~~処理終わり~~~~~~~~


'//-------------
'// ログ出力用
'// Call SysWriter("ログメッセージ", "test_log.txt")
'//-------------
Sub SysWriter(str,path)
 
	Dim fso, fi
	Set fso = CreateObject("Scripting.FileSystemObject")
	
	'ファイルを開く
	 'もしも存在しない場合には作成する
	Set fi = fso.OpenTextFile(path, 8, true)
	
	If Err.Number = 0 Then
		' ログ出力
		fi.WriteLine(Date() & " " & Time() & " - " & str)
		fi.Close
	Else
		' ログ出力用ファイルのオープンに失敗した場合
		WScript.Echo "ファイルオープンエラー: " & Err.Description
	End If
	
	Set fi = Nothing
 
End Sub


'//-----------------------------
'// プロセスが起動しているか調べる(あればtrue)
'//-----------------------------
Function ProcessExists(ProcessName)
	'返り値用 Array("プログラムの有無BOOL", Array( Process ID,,, (応答がなければ0) )) )
	Dim RetData(1)
	
	'プロセスチェック
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
		
	'応答なしのプログラムを探索する
		
		'タスクをすべてオブジェクトに入れる
		Dim objExec
		Set objExec = objWShell.Exec("WScript.exe " & ThisDir & "HideExec.vbs tasklist /fi ""status eq not responding""")
		objExec.StdIn.Close
		
		'出力結果
		Dim objStdOut
		Set objStdOut = objExec.StdOut
		
		'出力結果から探索
		Dim strLine, strTest
		Do Until objStdOut.AtEndOfStream
		
			strLine = objStdOut.ReadLine
			If (InStr(strLine, SERVER_PROG) <> 0) Then
			
				'該当プログラムが存在した場合
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
