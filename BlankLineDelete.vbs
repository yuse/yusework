'��s�폜�N ver1.0 2007/10/28
'�@�\�F�ΏۂP�t�@�C���̋�s���l�߂�

Const ForReading = 1
Const ForWriting = 2

Set objFSO = CreateObject("Scripting.FileSystemObject")
dim motofile
motofile = InputBox("�Ώۃt�@�C��(�t���p�X)���w�肵�ĉ������B")
chkmozi = Len(motofile)
if chkmozi <= 0 then
	Wscript.echo("�t�@�C���������͂���Ă��܂���B")
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
Wscript.echo("�u����Ƃ͏I�����܂����B")
Wscript.Quit