'GINGER_Description Rename file
'GINGER_$excelFilePath
'GINGER_$excelSheetName
'GINGER_$PARAM_NAME
'GINGER_$PARAM_VALUE
'GINGER_$GROUP_NAME


if WScript.Arguments.Count = 0 then
    WScript.Echo "Missing parameters"        
end if

Dim excelFilePath
Dim excelSheetName
Dim strValue
Dim SNO
Dim FoundCell,FoundCell_Empty  
Dim Data_Var
Dim PARAM_NAME
Dim PARAM_VALUE
Dim GROUP_NAME
Dim Required_File
Dim Data_Column
Dim intRow : intRow = 1

excelFilePath = WScript.Arguments(0)
excelSheetName = WScript.Arguments(1)
PARAM_NAME = WScript.Arguments(2)
PARAM_VALUE = WScript.Arguments(3)
'GROUP_NAME = trim(WScript.Arguments(4))

'############################################################
' Function name: FileExist
' Description:   
' Return value:  Success - True , Fail - False 
'#############################################################

Function FileExist(excelFilePath)
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	strFile = objFSO.FileExists(excelFilePath)
	
	If Not strFile Then
		WScript.Echo "File does not Exist at "& excelFilePath
		Exit Function
	End If

FileExist = True

End Function

'############################################################
' Function name: fWriteDataFile
' Description:   Use for Writing into Excel sheet
' Parameters:    None
' Return value:  Success - True , Fail - False     
'#############################################################

 Function fWriteDataFile(excelFilePath,excelSheetName)

	If FileExist(excelFilePath) = False Then
		WScript.Echo "File does not exist at "& excelFilePath		
		fWriteDataFile = "File does not exist at "& excelFilePath
		Exit Function
	End If
	
	Set objXls = CreateObject("Excel.Application")
	Set objWBook = objXls.Workbooks.Open(excelFilePath)
	Set objWSheet = objWBook.Worksheets(excelSheetName)

	If excelSheetName ="KEEP_REFER" then	
		intLength = Len(PARAM_NAME)
		PARAM_NAME = mid(PARAM_NAME, 2, intLength-1)
		cnt_rows = objWBook.WorkSheets("KEEP_REFER").usedRange.rows.count
		
  		Set FoundCell_VarName= objWBook.WorkSheets("KEEP_REFER").Range("A1:A65000").FIND(PARAM_NAME)	

			If Not FoundCell_VarName Is Nothing Then	
				objWBook.WorkSheets("KEEP_REFER").Range("B" & FoundCell_VarName.Row).Value = PARAM_VALUE
			Else	
				objWBook.WorkSheets("KEEP_REFER").Range("A" & cnt_rows+1).Value = PARAM_NAME
				objWBook.WorkSheets("KEEP_REFER").Range("B" & cnt_rows+1).Value = PARAM_VALUE'					
			End If	

			objWBook.Save
			objWBook.Close
			objXls.Quit
	
			Set objWSheet = Nothing
			Set objXls = Nothing
			Set objWBook=Nothing
			Exit Function

	 End if 
			
	Set FoundCell = objWSheet.Range("B1:B20000").Find("END")		
	If Not FoundCell Is Nothing Then
  		strValue =  FoundCell.Row
	End If
        
	Set FoundCell_Empty = objWSheet.Range("B1:B"&strValue).Find("")
		
On error resume Next

	
	Set FoundCell = objWSheet.Range("A" & FoundCell_Empty.Row & ":BZ" & FoundCell_Empty.Row).Find(PARAM_NAME,,,1)


	If Not FoundCell_Empty Is Nothing Then
		
		If PARAM_NAME="SKIP" AND PARAM_VALUE<>"N" Then
			
			objWSheet.Cells(FoundCell_Empty.Row,2).Value="X" 		
			objWSheet.Cells(FoundCell_Empty.Row+1,2).Value=PARAM_VALUE
		
		ElseIf PARAM_NAME="SKIP" AND PARAM_VALUE="N" Then
			
			Set FoundCell_FirstID = objWSheet.Range("A1:A"&strValue).Find("1")
			Set FoundCell_colNumber = objWSheet.Range("A1:ZZ"&strValue).Find("GROUP_NAME")			
			intCol = FoundCell_colNumber.Column
			intRow = FoundCell_FirstID.Row + 1
			For intRow = FoundCell_FirstID.Row + 1 to strValue-1			
			
				If trim(objWSheet.Cells(intRow,2)) = "" AND trim(objWSheet.Cells(intRow,intCol)) = GROUP_NAME Then
					objWSheet.Cells(intRow-1,2).Value = "X"
					objWSheet.Cells(intRow,2).Value = "N"
				End If
				intRow = intRow + 1
				If intRow >= strValue Then
					Exit For
				End If				
			Next				
			
		Else			 			
			objWSheet.Cells(FoundCell.Row+1,FoundCell.Column).Value = PARAM_VALUE 	
		End If       	  					   	
	        If instr(1,PARAM_NAME,"SCREEN_DATA_APP")>0 then 
 		  objWSheet.Cells(FoundCell.Row+1,FoundCell.Column)= Replace(objWSheet.Cells(FoundCell.Row+1,FoundCell.Column),"_"," ")
		End if			
	Else
		SNO = "none"
	End If
	
	objWBook.Save	
	objWBook.Close
	objXls.Quit
	
	Set objWSheet = Nothing
	Set objXls = Nothing
	Set FoundCell = Nothing
	set Copyrange = Nothing

fWriteDataFile = SNO
End Function

'*************************************************************************************************************************************************
 Function MakePropertyLibre(cName, uValue) 
    
  Dim oPropertyValue 
  Dim oSM 
	
  Set oSM = CreateObject("com.sun.star.ServiceManager")    
  Set oPropertyValue = oSM.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
  oPropertyValue.Name = cName
  oPropertyValue.Value = uValue
      
  Set MakePropertyLibre = oPropertyValue

End Function
'*************************************************************************************************************************************************

'############################################################

' Function name: fReadLibreFile
' Description:   Reading LibreOffice File 
' Parameters:  Excel File Path and Excel Sheet Name
'#############################################################
Function fWriteDataFile_Libre(excelFilePath,excelSheetName)

	If FileExist(excelFilePath) = False Then
		WScript.Echo "File Not Exist in "& excelFilePath		
		fWriteDataFile_Libre = "File Not Exist in "& excelFilePath
		Exit Function
	End If
     
	Set oSM = CreateObject("com.sun.star.ServiceManager")
	Set oDesk = oSM.createInstance("com.sun.star.frame.Desktop")
	Dim OpenPar(1)
	
	Set OpenPar(0) = MakePropertyLibre("Hidden", True)
	Set oDoc = oDesk.loadComponentFromURL("file:///" & excelFilePath, "_blank", 0, OpenPar)
	Set oSheet = oDoc.getSheets().getByName(excelSheetName)
		
	
	If excelSheetName ="KEEP_REFER" then			
		
		For i_EmptyRow1 = 1 to 500
						
						If oSheet.getCellByPosition(0,i_EmptyRow1).String = PARAM_NAME Then														
						
							oSheet.getCellByPosition(1,i_EmptyRow1).String =  PARAM_VALUE
							
							exit for
						End If	
					Next  		

		Required_File = Replace(excelFilePath, "\", "/")
		'msgbox Required_File 	
		oDoc.storeAsURL "file:///" & Required_File, OpenPar
			
		WScript.Sleep 5000
		oDoc.Close(True)	
		oDesk.terminate                                'Terminating the LibreOffice Object
		
		fWriteDataFile_Libre = SNO 
		Exit Function

	 End if 
		
		
	For i_End = 0 to 1000
		If oSheet.getCellByPosition(1, i_End).String = "END" Then
			FoundCell = i_End 
			Exit For
		End If
	Next	
	'msgbox FoundCell 	
	
	bFlagEmptyCellFound  = false
	bFlagEmptyDATAFound  = false
	'Find Empty Row
	For i_EmptyRow = 0 to FoundCell
		If oSheet.getCellByPosition(1, i_EmptyRow).String = "" Then
			bFlagEmptyCellFound  = true
			FoundCellEmpty = i_EmptyRow 
			Exit For
			
		ElseIf oSheet.getCellByPosition(1, i_EmptyRow).String = "END" Then
		
		 Exit Function
		End If
	Next	
	
	'msgbox FoundCellEmpty
	
	For i_EmptyColumn =0 to 1000
	
	If oSheet.getCellByPosition(i_EmptyColumn,FoundCellEmpty).String = PARAM_NAME Then
	
			'msgbox i_EmptyColumn
			bFlagEmptyDATAFound  = true
			Data_Column = i_EmptyColumn 
			Exit For
		End If
	Next	
	
	
	If FoundCellEmpty <> "" Then
		
		If PARAM_NAME="SKIP" AND PARAM_VALUE = "P" Then
			
			oSheet.getCellByPosition(1, FoundCellEmpty).String = "X"
			oSheet.getCellByPosition(1, FoundCellEmpty+1).String = PARAM_VALUE
			
		ElseIf PARAM_NAME="SKIP" AND PARAM_VALUE="F" Then
			FoundCell_colNumber = oSheet.getCellByPosition(7,FoundCellEmpty).String
			FoundCell_FirstID = oSheet.getCellByPosition(7, FoundCellEmpty+1).String
						
			'msgbox FoundCell_colNumber
			'msgbox FoundCell_FirstID
			
			If FoundCell_colNumber <> "" and FoundCell_FirstID <> "" Then
			
				oSheet.getCellByPosition(1, FoundCellEmpty).String = "X"
				oSheet.getCellByPosition(1, FoundCellEmpty+1).String = PARAM_VALUE
				
				For groupCount= FoundCellEmpty + 3 To FoundCell
						
						FoundCell_FirstID_New = oSheet.getCellByPosition(7, groupCount).String
						
					If (FoundCell_FirstID_New <> FoundCell_FirstID) = True Then		
						
						Exit For
					Else
						oSheet.getCellByPosition(1, groupCount-1).String = "X"
						oSheet.getCellByPosition(1, groupCount).String = "N"
					End If 
					
					groupCount = groupCount +1					
				Next
			End If	
		Else
			If bFlagEmptyDATAFound  <> false then
				oSheet.getCellByPosition(Data_Column, FoundCellEmpty+1).String = PARAM_VALUE
			End If		
		End If
	Else		
	End If
	

	Required_File = Replace(excelFilePath, "\", "/")
	'msgbox Required_File 	
	oDoc.storeAsURL "file:///" & Required_File, OpenPar
		
	WScript.Sleep 5000
	oDoc.Close(True)	
	oDesk.terminate                                'Terminating the LibreOffice Object
	
	fWriteDataFile_Libre = SNO 

 End Function
 '**********************************************************************************************************************************************************************************************
  
Call FileExist(excelFilePath)

If Right(excelFilePath,4) = ".ods" Then
	
	ReadData = fWriteDataFile_Libre(excelFilePath,excelSheetName)	
	Wscript.echo "~~~GINGER_RC_START~~~" 
	WScript.Echo "Outputvalue ="+ ReadData
	Wscript.echo "~~~GINGER_RC_END~~~"
	
	
 Else
	ReadData = fWriteDataFile(excelFilePath,excelSheetName)	
	Wscript.echo "~~~GINGER_RC_START~~~" 
	WScript.Echo "Outputvalue ="+ ReadData 	
	Wscript.echo "~~~GINGER_RC_END~~~"
	
 End If

Wscript.echo "~~~GINGER_RC_START~~~" 
WScript.Echo ReadData 
Wscript.echo "~~~GINGER_RC_END~~~"














