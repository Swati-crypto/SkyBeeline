'GINGER_Description Renamefile
'GINGER_$excelFilePath
'GINGER_$excelSheetName
'GINGER_$PARAM_NAME



'Option Explicit  'Line 10

if WScript.Arguments.Count = 0 then
    WScript.Echo "Missing parameters"        
end if

' Your code here
Dim excelFilePath
Dim excelSheetName
Dim strValue
Dim SNO
Dim FoundCell,FoundCell_Empty  
Dim Data_Var
Dim PARAM_NAME


excelFilePath = WScript.Arguments(0)
excelSheetName = WScript.Arguments(1)
PARAM_NAME = WScript.Arguments(2)




'############################################################

' Function name: FileExist

' Description:   

' Return value:  Success - True , Fail - False                           


'#############################################################

Function FileExist(excelFilePath)
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strFile = objFSO.FileExists(excelFilePath)
	
	If Not strFile Then
		WScript.Echo "File Not Exist in "& excelFilePath
		Exit Function
	End If

FileExist = True

End Function

'############################################################

' Function name: fReadDatafromDataFile

' Description:   

' Parameters:    None

' Return value:  Success - True , Fail - False                           



'#############################################################

 Function fReadDatafromDataFile(excelFilePath)

	If FileExist(excelFilePath) = False Then
		WScript.Echo "File Not Exist in "& excelFilePath		
		fReadDatafromDataFile = "File Not Exist in "& excelFilePath
		Exit Function
	End If

	
   Set objXls = CreateObject("Excel.Application")
   Set objWBook = objXls.Workbooks.Open(excelFilePath)
   Set objWSheet = objWBook.Worksheets(excelSheetName)

	
			
	Set FoundCell = objWSheet.Range("B1:B20000").Find("END")	
	
	If Not FoundCell Is Nothing Then
  		strValue =  FoundCell.Row          

	

	End If
		
	
        Set FoundCell_Empty = objWSheet.Range("B1:B" & strValue).Find("")
	
On Error resume Next	 

	Set FoundCell = objWSheet.Range("A" & FoundCell_Empty.Row & ":BZ" & FoundCell_Empty.Row).Find(PARAM_NAME, , , 1)
	

	If Not FoundCell Is Nothing Then

	If Not FoundCell_Empty Is Nothing Then

		'SNO = objWSheet.Cells(FoundCell.Row,1)
			       
		If PARAM_NAME="SKIP" Then
			
			objWSheet.Cells(FoundCell_Empty.Row,2).Value="X" 		
			objWSheet.Cells(FoundCell_Empty.Row+1,2).Value=PARAM_VALUE 		
		Else
			
			 PARAM_VALUE = objWSheet.Cells(FoundCell.Row+1,FoundCell.Column).Value
		
		End If 	  					   	
	       
	 
        
        	

	Else

		PARAM_VALUE = "none"

	End If
	End If
	


	objWBook.Save
	
	objWBook.Close

	objXls.Quit
	
	Set objWSheet = Nothing
	Set objXls = Nothing
	Set FoundCell = Nothing
	set Copyrange = Nothing

fReadDatafromDataFile = PARAM_VALUE 
 End Function

'*************************************************************************************************************************************************

Call FileExist(excelFilePath)

 ReadData = fReadDatafromDataFile(excelFilePath)



' output results to Ginger
Wscript.echo "~~~GINGER_RC_START~~~" 
WScript.Echo ReadData 
'Wscript.echo strVariable & "=" + ReadData 
Wscript.echo "~~~GINGER_RC_END~~~"














