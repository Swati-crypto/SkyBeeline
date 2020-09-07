'GINGER_$textData
textData=WScript.Arguments(0)
'MsgBox textData
'Data=replace(Data,"{","")
'Data=replace(Data,"}","")
Data=replace(textData, chr(34), "")
Data=replace(Data,"{","")
Data=replace(Data,"}","")
'MsgBox Data
WScript.Echo Data