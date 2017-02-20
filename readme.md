CheckRerun
====

概要

プロセスを監視して，指定したプログラムが実行されていなければ，指定したプログラムを実行してくれるスクリプト．

自身のbatやVBScriptの勉強を兼ねて作成しました．

[.exe]のフリーソフトを導入することなく，簡単にプログラムの監視を行ってくれるスクリプトを目指して作成しました．

個人で立てているサーバ用プログラムが突然終了していたり，原因不明の応答なし現象時に自動で復帰する様に設計しています．

例えば，MODを導入しているゲーム用サーバプログラムを自動で復帰する状況を想定しています．

#### 現状の修正点
 応答なし状態で，実際に実行してくれるかどうかのデバッグ環境が，実際に想定されうる状況が乖離しているため，
 応答なし状態での動作が未サポート状態です．
 
 HideExec.vbsを書き換えて，CheckRerun.vbsに統合予定です．

## 使い方

 1. CheckRerun.vbsをメモ帳で開いて，以下の部分を環境に合わせてプログラムの指定を行う様に修正してください．
 
    ' ここに監視を行うserverプログラムの"プロセス名"を入力する
    Const SERVER_PROG = "cmd.exe"
    
    ' ここに実行を行うserverプログラムの"場所"を入力する
    Const SERVER_DIR = "C:\Windows\System32"
    
    ' ここに実行を行うserverプログラムの"名前"を入力する
    Const SERVER_EXEC = "notepad.exe"
    
    ' 起動していなかった時に、再起動した日時を出力してくれるファイル
    Const LOG_PROG = "process_test.txt"
 
 上記箇所を修正後，指定したCheckRerun.vbsをダブルクリックやコマンドラインから実行することで，指定したプログラムが実行されるか確認を行ってください．
 
 なお，実行はアクティブ状態を変更せず，最小化状態で実行されるます．
 そのため，ディフォルトの状態であると，知らず知らずのうちに相当な数の[メモ帳]が実行されてしまいます．
 
 この例のように，監視を行うプロセス名と，それが応答していない場合に実行を行うプログラムを変更することが出来ます．
 
 ここで，指定するプログラムを例えば*.bat*ファイルを指定することも可能です．
 
 しかし，それが実行したことでプロセスの監視対象が実行されたかどうかの確認は行わないため，注意してください．
 
 プログラムが実行されていない，もしくは，応答なし状態時に，指定したログファイルに日時と実行したプログラム名を出力します．
 
 2. server_reg.batをダブルクリックなどをして実行することで，一定時間毎で実行してくれるタスクスケジューラにCheckRerunという名前で登録されます．
 
 ディフォルトでは5分間隔で実行されます．
 
 確認が行われる間隔はserver_reg.batをメモ帳で開くことで変更を行うことができます．最短単位は1分毎です．
 
 3. もし，一定時間ごとに実行される動作を終了させたい場合はserver_reg_uninstall.batを実行してください．

## 各スクリプトの説明
 
*CheckRerun.vbs
　プログラムが実行されてるか(プロセス一覧の中に指定したプログラムがあるか)
　を確認して、無ければそのプログラムを実行してくれるプログラム
　同時に、プロセスが応答なしであれば、そのプロセスを強制終了させてから、
　プロセスを実行してくれるプログラム。
　
　ディフォルトではprocess_test.txtにログを出力してくれます。
　起動されていなかった場合、応答なしの場合にのみ、日時と共にログ出力を行います。
　
　実行されるプログラムはウィンドウのアクティブ状態を変更せず、最小化状態で実行されます。

*HideExec.vbs
　プログラムを裏で実行するために相対的に参照するプログラム。
　後にCheckRerun.vbsに統合予定。
　thank you ["Windows Script Programming"](http://scripting.cocolog-nifty.com/blog/2008/08/wscriptshellexe_7621.html "WScript.ShellのExec()で、コンソールアプリを非表示で実行するラッパー")!

*server_reg.bat
　指定した時間間隔でプログラムを実行してくれるプログラムに登録する為のプログラム
　CheckRerunという名前で登録します．

*server_reg_uninstall.bat
　server_reg.batで登録したプログラムを削除してくれるプログラム

*process_test.txt
 ログ出力用テキストファイル．
 
## Usage
 
 
## Install
 
 
## Licence
 まだだよ．
 完全にオープンソースで複製自由に

## 何かあれば(Author)
　[Twitter@meto4d](https://twitter.com/meto4d)
 [meto4d](https://github.com/meto4d)
