'空行削除君 ver1.1 2015/10/15
'機能：指定されたファイルの空行を詰めて保存する

Const ForReading = 1
Const ForWriting = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")
dim motofile
motofile = InputBox("対象ファイル(フルパス)を指定して下さい。")
chkmozi = Len(motofile)
if chkmozi <= 0 then
	Wscript.echo("ファイル名が入力されていません。")
	Wscript.Quit
end if


Set objFile = objFSO.OpenTextFile(motofile, ForReading)

Do Until objFile.AtEndOfStream
    strLine = objFile.Readline
    strLine = Trim(strLine)
    If Len(strLine) > 0 Then
        strNewContents = strNewContents & strLine & vbCrLf
    End If
Loop

objFile.Close

Set objFile = objFSO.OpenTextFile(motofile, ForWriting)
objFile.Write strNewContents
objFile.Close
Wscript.echo("置換作業は終了しました。")
Wscript.Quit