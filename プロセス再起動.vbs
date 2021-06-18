Option Explicit
'On Error Resume Next

' ------------- 定数宣言（いじるならココ！！） ---------------

' 再起動させるスイッチファイル（フルパスで。ex.「C:\work\reboot」）
Const CST_CHECK_FILE = "C:\work\reboot"

' 起動させるプログラム名（フルパスで。ex.「%windir%\system32\notepad.exe」)
Const CST_CMD_LINE = "%windir%\system32\notepad.exe"


' ------------- 変数宣言 ---------------

Dim objFSO		' FileSystemObject
Dim bln_Switch	' スイッチファイルの有無
Dim objProcList	' プロセス一覧
Dim objProcess	' プロセス情報
Dim objWshShell	' WshShell オブジェクト
Dim objFile		' 書き出しファイル
Dim strLog		' ログファイル名
Dim strProgName	' 起動プログラムのファイル名

' ■スイッチ用のファイルを検索する
' ■スイッチ用のファイルがないと、その場で終了

	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

	If Err.Number = 0 Then
		If objFSO.FileExists(CST_CHECK_FILE) = True Then
			Call Output_log("ファイルあり: " & CST_CHECK_FILE)
			' ファイルを削除する
			objFSO.DeleteFile CST_CHECK_FILE, True
			If Err.Number <> 0 Then
				Call Output_log("エラー: " & Err.Description)
			End If
		else
			WScript.Quit
		End If
	Else
		Call Output_log("エラー: " & Err.Description)
	End If


' ■起動プログラムがプロセスにあれば、terminate（kill）する。

	strProgName = objFSO.GetFileName(CST_CMD_LINE)

	Set objProcList = GetObject("winmgmts:").InstancesOf("win32_process")
	For Each objProcess In objProcList
		If LCase(objProcess.Name) = LCase(strProgName) Then
			' プロセスを強制終了する
			objProcess.Terminate
			Call Output_log(strProgName & " を強制終了しました。")
			' 5秒待つ
			WScript.Sleep 5000
			If Err.Number <> 0 Then
				Call Output_log("エラー: " & Err.Description)
			End If
		End If
	Next



' ■再度起動する

	Set objWshShell = WScript.CreateObject("WScript.Shell")
	If Err.Number = 0 Then
		' コマンドを実行する
		objWshShell.Exec(CST_CMD_LINE)
		If Err.Number = 0 Then
			Call Output_log(CST_CMD_LINE & " を起動しました。")
		Else
			Call Output_log("エラー: " & Err.Description)
		End If
	Else
		Call Output_log("エラー: " & Err.Description)
	End If


' ■終了処理
	Set objFSO = Nothing
	Set objProcList = Nothing
	Set objWshShell = Nothing

	WScript.Quit

' ■【関数】エラー出力
	Sub Output_log(strMsg)

		' ログファイル名を作成（プログラム名+log）
		strLog = objFSO.GetParentFolderName( WScript.ScriptFullName) & "\" & _
				Left(WScript.ScriptName, Len(WScript.ScriptName)-4) & ".log"

		' ログファイルを開き、書き込み、閉じる
		Set objFile = objFSO.OpenTextFile(strLog, 8, True)
		objFile.WriteLine(Now() & " " & strMSG)
		objFile.Close
		Set objFile = Nothing

	End Sub



