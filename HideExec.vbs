'//--------------------
'//    http://scripting.cocolog-nifty.com/blog/2008/08/wscriptshellexe_7621.html
'//    WScript.ShellのExec()で、コンソールアプリを非表示で実行するラッパー
'//--------------------
Option Explicit
Dim fso
Dim wStdIn
Dim wStdOut
Dim wStdErr
Dim TempPath
Dim pStdIn
Dim pStdOut
Dim pStdErr
Dim StdIn
Dim StdOut
Dim StdErr
Dim wShell
Dim Command
Dim arg

Set fso=CreateObject("Scripting.FileSystemObject")
Set wStdIn=fso.GetStandardStream(0)
Set wStdOut=fso.GetStandardStream(1)
Set wStdErr=fso.GetStandardStream(2)
TempPath=fso.GetSpecialFolder(2).Path
pStdIn=fso.BuildPath(TempPath,fso.GetTempName())
pStdOut=fso.BuildPath(TempPath,fso.GetTempName())
pStdErr=fso.BuildPath(TempPath,fso.GetTempName())
Set StdIn=fso.CreateTextFile(pStdIn)
Sub Try
Do While Not wStdIn.AtEndOfStream
  StdIn.Write wStdIn.Read(1)
Loop
End Sub
Sub Catch
On Error Resume Next
Try
End Sub
Catch
StdIn.Close
Set wShell=CreateObject("WScript.Shell")
Command=Array()
For Each arg In WScript.Arguments
  If InStr(arg," ") Then arg="""" & arg & """"
  ReDim Preserve Command(UBound(Command)+1)
  Command(UBound(Command))=arg
Next
Command=Join(Command)
Command=Replace(Command,"`","""")
Call CreateObject("WScript.Shell").Run("CMD.EXE /S /C <""" & pStdIn & """ "& Command & " 1>""" & pStdOut & """ 2>""" & pStdErr & """""",0,False)
Do While Not fso.FileExists(pStdOut)
  WScript.Sleep 1000
Loop
Set StdOut=fso.OpenTextFile(pStdOut)
Do While Not fso.FileExists(pStdErr)
  WScript.Sleep 1000
Loop
Set StdErr=fso.OpenTextFile(pStdErr)
Do
  Do While Not StdOut.AtEndOfStream
    On Error Resume Next
    wStdOut.Write StdOut.Read(1)
    On Error GoTo 0
  Loop
  Do While Not StdErr.AtEndOfStream
    On Error Resume Next
    wStdErr.Write StdErr.Read(1)
    On Error GoTo 0
  Loop
  If AtEndOfStream(pStdOut) And AtEndOfStream(pStdErr) And StdOut.AtEndOfStream And StdErr.AtEndOfStream Then Exit Do
  WScript.Sleep 1000
Loop
StdOut.Close
StdErr.Close
fso.DeleteFile(pStdIn)
fso.DeleteFile(pStdOut)
fso.DeleteFile(pStdErr)

Function AtEndOfStream(Path)
On Error Resume Next
Call fso.OpenTextFile(Path,8)
AtEndOfStream=CBool(Err.Number=0)
End Function