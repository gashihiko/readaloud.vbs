'�ǂݏグ�v���O�������I�������邽�߂̎d�g��(
'�Q�l: https://okwave.jp/qa/q7861047.html
cnt = 0
for each p in GetObject("winmgmts:{impersonationLevel=impersonate}"). _
	ExecQuery("select * from Win32_Process where Name='wscript.exe'")
	if instr(p.CommandLine,"""" & WScript.ScriptFullName & """")>0 then
		cnt = cnt+1
	end if
next

Sub ReadTerminate
	'�N�����ꂽ���Ƀv���Z�X���I�΂��d�l�炵���B�Ȃ̂Ŏ������g�͍Ō�ɏI�������B
	For Each p in GetObject("winmgmts:{impersonationLevel=impersonate}"). _
		ExecQuery("select * from Win32_Process where Name='wscript.exe'")
		If InStr(p.CommandLine, "readaloud.vbs") Then
			p.Terminate
		End If
	Next
End Sub

If cnt = 1 Then '�������I�������邽�߂̃v���Z�X���N��
	Set WshShell = WScript.CreateObject("WScript.Shell")
	Call WshShell.Run("wscript " & """" & WScript.ScriptFullName & """", 0, False)
Else '�������N���������̂��I�������邽�߂̃v���Z�X�Ɖ����B
	ans = Msgbox ("�I��!?", vbOKOnly, "readaloud.vbs")
	If ans = vbOK Then
		Call ReadTerminate
	End If
End If
'�ǂݏグ�v���O�������I�������邽�߂̎d�g��)

strFileText = ""
For Each fp in WScript.Arguments '�h���b�v���ꂽ�t�@�C�������ɓǂݏグ�邽�߂̃e�L�X�g���쐬�B�G���R�[�f�B���O��������B
	'�Q�l: https://replication.hatenablog.com/entry/20080218/1203352273
	With CreateObject("ADODB.Stream")
		.Open
		.Charset = "_autodetect_all"
		.LoadFromFile(fp)
		strFileText = strFileText & .ReadText
		.Close
	End With
Next

Set objFileToRead = Nothing

for i=1 to 9
	Readaloud(strFileText)
next

Sub Readaloud(text) 'Windows�g�ݍ��݋@�\�𗘗p�����ǂݏグ�T�u�v���V�[�W���B���{��ǂ݁E�p��ǂݗ��Ή��B
    Set Regx = CreateObject("VBScript.RegExp")
    With Regx
		.Pattern = "^[a-zA-Z0-9!-/:-@\[-`{-~\s]+$"
        .IgnoreCase = True
        .Global = True
    End With
    IsAscii = Regx.Test(text) '���K�\����ASCII�݂̂̕����񂩂ǂ����𔻕�

    With CreateObject("SAPI.SpVoice")
		.Rate = 0 'Values for the Rate property range from -10 to 10
        If IsAscii Then
            Set .Voice = .GetVoices.Item(1) '�p��ǂ݃��[�h�ɐ؂�ւ�
        Else
            Set .Voice = .GetVoices.Item(0) '���{��ǂ݃��[�h�ɐ؂�ւ�
        End If
        .Speak text
    End With
End Sub
