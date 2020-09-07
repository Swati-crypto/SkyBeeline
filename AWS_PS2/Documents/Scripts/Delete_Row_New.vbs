
'GINGER_Description Renamefile
'GINGER_$excelFilePath
'GINGER_$excelSheetName
'GINGER_$PARAM_NAME
'GINGER_$PARAM_VALUE
'GINGER_$BUSINESS_FLOW_NAME

'Option Explicit  'Line 10
'Closee Opened Excel


' Your code here
Dim excelFilePath
Dim excelSheetName
Dim strValue
Dim SNO
Dim FoundCell,FoundCell_Empty  
Dim Data_Var
Dim PARAM_NAME
Dim PARAM_VALUE
Dim BUSINESS_FLOW_NAME

excelFilePath = WScript.Arguments(0)
excelSheetName = WScript.Arguments(1)
PARAM_NAME = WScript.Arguments(2)
PARAM_VALUE = WScript.Arguments(3)
BUSINESS_FLOW_NAME = WScript.Arguments(4)

'===========================================
'Code Updated By Subodh Singh
'Date- 18-July-2018
'Description- Handle relative path for excel 
'===========================================
If Instr(excelFilePath,"rel::") = 1 Then 
	Set objFSObject = CreateObject("Scripting.FileSystemObject")
	excelFilePath = replace(excelFilePath,"rel::","")
	vbsFullName = Wscript.ScriptFullName
	vbsFile = objFSObject.GetFile(vbsFullName)
	scriptsFullPath = objFSObject.GetParentFolderName(vbsFile) 
	documentsFullPath=objFSObject.GetParentFolderName(scriptsFullPath)
	excelFilePath = documentsFullPath&"\"&excelFilePath
End If
'===========================================
'Code Change End
'===========================================

excelFilePath1 = Replace(excelFilePath,"\","/")
excelFilePath1 = "file:///"&excelFilePath1



'############################################################

' Function name: FileExist

' Description:   

' Return value:  Success - True , Fail - False                           


'#############################################################

Function MakePropertyValue(cName, uValue) 
    
  Dim oPropertyValue 
  Dim oSM 
	
  Set oSM = CreateObject("com.sun.star.ServiceManager")    
  Set oPropertyValue = oSM.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
  oPropertyValue.Name = cName
  oPropertyValue.Value = uValue
      
  Set MakePropertyValue = oPropertyValue

End Function

Function fWriteDataToFile(excelFilePath)

	Dim OpenPar(2)
	
	If FileExist(excelFilePath) = False Then
		WScript.Echo "File Not Exist in "& excelFilePath		
		fReadDatafromDataFile = "File Not Exist in "& excelFilePath
		Exit Function
	End If

	
   'Set objXls = CreateObject("Excel.Application")
   'Set objWBook = objXls.Workbooks.Open(excelFilePath)
   'Set objWSheet = objWBook.Worksheets(excelSheetName)

	Set oSM = CreateObject("com.sun.star.ServiceManager")
	Set oDesk = oSM.createInstance("com.sun.star.frame.Desktop")
	Dim arg()
	
	'Set OpenPar(0) = MakePropertyValue("ReadOnly", True)
	'Set OpenPar(1) = MakePropertyValue("Password", "secret")
	Set OpenPar(1) = MakePropertyValue("Hidden", True)
	Set wb = oDesk.loadComponentFromURL(excelFilePath1, "_blank", 0, OpenPar)
	Set oSheet = wb.CurrentController.ActiveSheet
	
	'msgbox oSheet.getCellByPosition(1, 2).String
	
		
	'Get Row Number where value is END 
	For i = 0 to 1000
		If oSheet.getCellByPosition(1, i).String = "END" Then
			FoundCell = i 
			Exit For
		End If
	Next	
	
	'msgbox FoundCell
	
	bFlagEmptyCellFound  = false
	'Find Empty Row
	For i = 0 to FoundCell
		If oSheet.getCellByPosition(1, i).String = "" Then
			bFlagEmptyCellFound  = true
			FoundCellEmpty = i 
			Exit For
		End If
	Next	
	
	If bFlagEmptyCellFound Then
		'msgbox FoundCellEmpty
	Else
		'msgbox "Empty cell Not Found"
	End If 
	
	'Find Value for column as PARAM_NAME
	For i = 2 to 50
		If oSheet.getCellByPosition(i, FoundCellEmpty).String = PARAM_NAME Then
			FoundColumnNum = i 
		
		If PARAM_NAME="SKIP" Then
			
			oSheet.getCellByPosition(2, FoundCellEmpty).setString("X")
			'objWSheet.Cells(FoundCell_Empty.Row,2).Value="X" 		
			'objWSheet.Cells(FoundCell_Empty.Row+1,2).Value=PARAM_VALUE 	
			oSheet.getCellByPosition(2, FoundCellEmpty+1).setString(PARAM_VALUE)
		Else
			 			
			 oSheet.getCellByPosition(i, FoundCellEmpty+1).setString(PARAM_VALUE)
	
		End If
			
			oSheet.getCellByPosition(i, FoundCellEmpty+1).setString(PARAM_VALUE)
			'msgbox PARAM_VALUE
			Exit For
		End If
	Next
	
	wb.Close(True)
	Set wb = Nothing
	Set oSM = Nothing
	Set oSheet = Nothing


fWriteDataToFile = PARAM_VALUE 

 End Function

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

		'Check and close if excel is opened
	Set objExcel = GetObject(excelFilePath).Application 'Use this syntax for multiple instances of Excel.
	'msgbox excelFilePath
	If (not objExcel.ActiveWorkbook is nothing) then
		'msgbox "ActiveWorkbook is: " & objExcel.ActiveWorkbook.Name
			If Instr(excelFilePath,objExcel.ActiveWorkbook.name)>0 then
				'msgbox "Closing Workbook " & excelFilePath 
				objExcel.ActiveWorkbook.Saved = True
				objExcel.DisplayAlerts = False

				'objWorkbook.Close False 
				objExcel.ActiveWorkbook.Close
				'objExcel.Close	
			End if
		Else
			'wscript.echo "Open Workbook " & strName & " Not Found"
	End if

	set objExcel=nothing

   Set objXls = CreateObject("Excel.Application")
   Set objWBook = objXls.Workbooks.Open(excelFilePath)
   Set objWSheet = objWBook.Worksheets(excelSheetName)
	
			
	Set FoundCell = objWSheet.Range("B1:B30000").Find("END")
	
	If Not FoundCell Is Nothing Then
  		strValue =  FoundCell.Row-1
	End If
	
	'msgbox FoundCell
	'msgbox strValue

	
	If PARAM_NAME="DELSKIP" Then
	
	'Modified by shridhar to delete the skip row to run the flow BF wise
	
		    For i_EmptyRow1 = 1 to strValue		
				'msgbox i_EmptyRow1
				
				data = objWSheet.Cells(i_EmptyRow1,3).Value
				'msgbox data
				
							'msgbox BUSINESS_FLOW_NAME
							
				If data = BUSINESS_FLOW_NAME Then
				objWSheet.Cells(i_EmptyRow1,2).Value=""
				objWSheet.Cells(i_EmptyRow1-1,2).Value=""
				
				'msgbox i_EmptyRow1
				'msgbox FoundCell
				
			
			 'oSheet.getCellByPosition(1,i_EmptyRow1).String =""
			 
			 If i_EmptyRow1 = strValue-1 Then
				exit for
			 End If	
			  End If	
		
			Next	
	
			'data = objWSheet.Cells(1,2).Value
			'msgbox data
		
						
			'objWSheet.Range("B2:B"&strValue).ClearContents
			'objWSheet.Cells(FoundCell_Empty.Row,2).Value="X" 		
			'objWSheet.Cells(FoundCell_Empty.Row+1,2).Value=PARAM_VALUE
				objWBook.Save
				objWBook.Close
				objXls.Quit
			Exit Function		
	End If
	
	
    Set FoundCell_Empty = objWSheet.Range("B1:B"&strValue).Find("")

	Set FoundCell = objWSheet.Range("A" & FoundCell_Empty.Row & ":BZ" & FoundCell_Empty.Row).Find(PARAM_NAME,,,1)
				
	If Not FoundCell_Empty Is Nothing Then

	'SNO = objWSheet.Cells(FoundCell.Row,1)
	       
		If PARAM_NAME="SKIP" Then
				
			objWSheet.Cells(FoundCell_Empty.Row,2).Value="X" 		
			objWSheet.Cells(FoundCell_Empty.Row+1,2).Value=PARAM_VALUE 		
				
		Else If Not FoundCell Is Nothing Then
							
			 objWSheet.Cells(FoundCell.Row+1,FoundCell.Column).Value = PARAM_VALUE 
			
		Else

			SNO = "none"

		End If

	'End If
		
	end if
       	  					   	
	If instr(1,PARAM_NAME,"SCREEN_DATA_APP")>0 then

	  objWSheet.Cells(FoundCell.Row+1,FoundCell.Column)= Replace(objWSheet.Cells(FoundCell.Row+1,FoundCell.Column),"_"," ")
	End if	  	
	             	
	Else

		SNO = "none"
		
	End If
	'objWSheet.Save	
	'objWSheet.Close	
	
	objWBook.Save	
	objWBook.Close	
	objXls.Quit
		
	Set objWSheet = Nothing
	Set objXls = Nothing
	Set FoundCell = Nothing
	set Copyrange = Nothing
	
fReadDatafromDataFile = SNO
 End Function
 
 
'*******************************

'############################################################

' Function name: fBlankRowFile_Libre

' Description:   

' Parameters:    None                 



'#############################################################


Function fBlankRowFile_Libre(excelFilePath,excelSheetName)

	If FileExist(excelFilePath) = False Then
		WScript.Echo "File Not Exist in "& excelFilePath		
		fBlankRowFile_Libre = "File Not Exist in "& excelFilePath
		Exit Function
	End If
     
	Set oSM = CreateObject("com.sun.star.ServiceManager")
	Set oDesk = oSM.createInstance("com.sun.star.frame.Desktop")
	Dim OpenPar(1)
	
	Set OpenPar(0) = MakePropertyValue("Hidden", True)
	Set oDoc = oDesk.loadComponentFromURL(excelFilePath1, "_blank", 0, OpenPar)
	Set oSheet = oDoc.getSheets().getByName(excelSheetName)
	
	For i_End = 0 to 1000
		If oSheet.getCellByPosition(1, i_End).String = "END" Then
			FoundCell = i_End 
			'msgbox FoundCell
			Exit For
		End If
	Next		
	

		
	If PARAM_NAME="DELSKIP" Then
			
			'Modified by shridhar to delete the skip row to run the flow BF wise
			
			For i_EmptyRow1 = 1 to FoundCell	

			data = oSheet.getCellByPosition(2,i_EmptyRow1).String	
			
			If data = BUSINESS_FLOW_NAME Then
			oSheet.getCellByPosition(1,i_EmptyRow1).String =""
			oSheet.getCellByPosition(1,i_EmptyRow1-1).String =""
					 
			 If i_EmptyRow1 = FoundCell-1 Then
				exit for
			 End If	
			 End If
		
			Next			
			
			Required_File = Replace(excelFilePath, "\", "/")
		'msgbox Required_File 	
		oDoc.storeAsURL "file:///" & Required_File, OpenPar
			
		WScript.Sleep 5000
		oDoc.Close(True)	
		oDesk.terminate                              'Terminating the LibreOffice Object
				
		Exit Function

	 End if 
	
	

 End Function
 

'*************************************************************************************************************************************************

Call FileExist(excelFilePath)

 'ReadData = fReadDatafromDataFile(excelFilePath)

If Right(excelFilePath,4) = ".ods" Then
 
	ReadData = fBlankRowFile_Libre(excelFilePath,excelSheetName)
 Else
	ReadData = fReadDatafromDataFile(excelFilePath)
 End If

' output results to Ginger
Wscript.echo "~~~GINGER_RC_START~~~" 
WScript.Echo ReadData 
'Wscript.echo strVariable & "=" + ReadData 
Wscript.echo "~~~GINGER_RC_END~~~"














