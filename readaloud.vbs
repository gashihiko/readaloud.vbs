'読み上げプログラムを終了させるための仕組み(
'参考: https://okwave.jp/qa/q7861047.html
cnt = 0
for each p in GetObject("winmgmts:{impersonationLevel=impersonate}"). _
	ExecQuery("select * from Win32_Process where Name='wscript.exe'")
	if instr(p.CommandLine,"""" & WScript.ScriptFullName & """")>0 then
		cnt = cnt+1
	end if
next

Sub ReadTerminate
	'起動された順にプロセスが選ばれる仕様らしい。なので自分自身は最後に終了される。
	For Each p in GetObject("winmgmts:{impersonationLevel=impersonate}"). _
		ExecQuery("select * from Win32_Process where Name='wscript.exe'")
		If InStr(p.CommandLine, "readaloud.vbs") Then
			p.Terminate
		End If
	Next
End Sub

If cnt = 1 Then '自分を終了させるためのプロセスを起動
	Set WshShell = WScript.CreateObject("WScript.Shell")
	Call WshShell.Run("wscript " & """" & WScript.ScriptFullName & """", 0, False)
Else '自分を起動したものを終了させるためのプロセスと化す。
	ans = Msgbox ("終了!?", vbOKOnly, "readaloud.vbs")
	If ans = vbOK Then
		Call ReadTerminate
	End If
End If
'読み上げプログラムを終了させるための仕組み)

strFileText = ""
For Each fp in WScript.Arguments 'ドロップされたファイルを順に読み上げるためのテキストを作成。エンコーディング自動判定。
	'参考: https://replication.hatenablog.com/entry/20080218/1203352273
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

Sub Readaloud(text) 'Windows組み込み機能を利用した読み上げサブプロシージャ。日本語読み・英語読み両対応。
    Set Regx = CreateObject("VBScript.RegExp")
    With Regx
		.Pattern = "^[a-zA-Z0-9!-/:-@\[-`{-~\s]+$"
        .IgnoreCase = True
        .Global = True
    End With
    IsAscii = Regx.Test(text) '正規表現でASCIIのみの文字列かどうかを判別

    With CreateObject("SAPI.SpVoice")
		.Rate = 0 'Values for the Rate property range from -10 to 10
        If IsAscii Then
            Set .Voice = .GetVoices.Item(1) '英語読みモードに切り替え
        Else
            Set .Voice = .GetVoices.Item(0) '日本語読みモードに切り替え
        End If
        .Speak text
    End With
End Sub
