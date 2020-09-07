'GINGER_$textData
textData=WScript.Arguments(0)

Set fso = CreateObject("Scripting.FileSystemObject")
FilePath =fso.GetAbsolutePathName("..\JsonRequestFiles\..")

Set fh = fso.OpenTextFile(FilePath & "\" & "JsonRequestFiles" & "\" & "JsonRequestFile.txt")
strin = Split(fh.readall, vbNewLine)
a= strin(fh.Line - textData)
Data=replace(a, Chr(34), "")
Data=replace(Data, Chr(32), "")
Data=replace(Data, Chr(44), "")
wscript.echo Data