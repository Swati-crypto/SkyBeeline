'GINGER_$textData
textData=WScript.Arguments(0)
SET FS = CreateObject("Scripting.FileSystemObject")
SET StdOut = FS.GetStandardStream(1)
Wscript.echo "~~~GINGER_RC_START~~~"
StdOut.Write("Outputvalue =")
Wscript.StdOut.Write(textData)
Wscript.echo "~~~GINGER_RC_END~~~"
'WScript.Echo textData