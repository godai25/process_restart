Option Explicit
'On Error Resume Next

' ------------- �萔�錾�i������Ȃ�R�R�I�I�j ---------------

' �ċN��������X�C�b�`�t�@�C���i�t���p�X�ŁBex.�uC:\work\reboot�v�j
Const CST_CHECK_FILE = "C:\work\reboot"

' �N��������v���O�������i�t���p�X�ŁBex.�u%windir%\system32\notepad.exe�v)
Const CST_CMD_LINE = "%windir%\system32\notepad.exe"


' ------------- �ϐ��錾 ---------------

Dim objFSO		' FileSystemObject
Dim bln_Switch	' �X�C�b�`�t�@�C���̗L��
Dim objProcList	' �v���Z�X�ꗗ
Dim objProcess	' �v���Z�X���
Dim objWshShell	' WshShell �I�u�W�F�N�g
Dim objFile		' �����o���t�@�C��
Dim strLog		' ���O�t�@�C����
Dim strProgName	' �N���v���O�����̃t�@�C����

' ���X�C�b�`�p�̃t�@�C������������
' ���X�C�b�`�p�̃t�@�C�����Ȃ��ƁA���̏�ŏI��

	Set objFSO = WScript.CreateObject("Scripting.FileSystemObject")

	If Err.Number = 0 Then
		If objFSO.FileExists(CST_CHECK_FILE) = True Then
			Call Output_log("�t�@�C������: " & CST_CHECK_FILE)
			' �t�@�C�����폜����
			objFSO.DeleteFile CST_CHECK_FILE, True
			If Err.Number <> 0 Then
				Call Output_log("�G���[: " & Err.Description)
			End If
		else
			WScript.Quit
		End If
	Else
		Call Output_log("�G���[: " & Err.Description)
	End If


' ���N���v���O�������v���Z�X�ɂ���΁Aterminate�ikill�j����B

	strProgName = objFSO.GetFileName(CST_CMD_LINE)

	Set objProcList = GetObject("winmgmts:").InstancesOf("win32_process")
	For Each objProcess In objProcList
		If LCase(objProcess.Name) = LCase(strProgName) Then
			' �v���Z�X�������I������
			objProcess.Terminate
			Call Output_log(strProgName & " �������I�����܂����B")
			' 5�b�҂�
			WScript.Sleep 5000
			If Err.Number <> 0 Then
				Call Output_log("�G���[: " & Err.Description)
			End If
		End If
	Next



' ���ēx�N������

	Set objWshShell = WScript.CreateObject("WScript.Shell")
	If Err.Number = 0 Then
		' �R�}���h�����s����
		objWshShell.Exec(CST_CMD_LINE)
		If Err.Number = 0 Then
			Call Output_log(CST_CMD_LINE & " ���N�����܂����B")
		Else
			Call Output_log("�G���[: " & Err.Description)
		End If
	Else
		Call Output_log("�G���[: " & Err.Description)
	End If


' ���I������
	Set objFSO = Nothing
	Set objProcList = Nothing
	Set objWshShell = Nothing

	WScript.Quit

' ���y�֐��z�G���[�o��
	Sub Output_log(strMsg)

		' ���O�t�@�C�������쐬�i�v���O������+log�j
		strLog = objFSO.GetParentFolderName( WScript.ScriptFullName) & "\" & _
				Left(WScript.ScriptName, Len(WScript.ScriptName)-4) & ".log"

		' ���O�t�@�C�����J���A�������݁A����
		Set objFile = objFSO.OpenTextFile(strLog, 8, True)
		objFile.WriteLine(Now() & " " & strMSG)
		objFile.Close
		Set objFile = Nothing

	End Sub



