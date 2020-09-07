'GINGER_Description Rename file
'GINGER_$excelFilePath
'GINGER_$excelSheetName

if WScript.Arguments.Count = 0 then
    WScript.Echo "Missing parameters"        
end if

Dim excelFilePath
Dim excelSheetName
Dim SNO
Dim PARAM_Cell_COLUMN
Dim END_ROW
DIM Found_Empty_Row
Dim strValue
Dim Data_Var
Dim Found_Last_Column
	
excelFilePath = WScript.Arguments(0)
excelSheetName = WScript.Arguments(1)

'############################################################
' Description:   New object creation for "StdOut.Write" 
' Modification date:  11 Dec 2018                          

'#############################################################

SET FS = CreateObject("Scripting.FileSystemObject")
SET StdOut = FS.GetStandardStream(1)

'############################################################
' Function name: FileExist
' Description:  To check input file present on specify path  
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

' Function name: MakePropertyLibre
' Description:   Use for Making Libre office file Read only and hidden 
' Parameters:  Name and Value
'#############################################################

Function MakePropertyLibre(cName, uValue) 
    
  Dim oPropertyValue 
  Dim oSM 
	
  Set oSM = CreateObject("com.sun.star.ServiceManager")    
  Set oPropertyValue = oSM.Bridge_GetStruct("com.sun.star.beans.PropertyValue")
  oPropertyValue.Name = cName
  oPropertyValue.Value = uValue      
  Set MakePropertyLibre = oPropertyValue

End Function


'############################################################
' Function name: fReadDatafromDataFile
' Description:   
' Parameters:    None
' Return value:  Success - True , Fail - False          
'#############################################################

 Function fReadExcelFile(excelFilePath,excelSheetName)

	If FileExist(excelFilePath) = False Then 				' Calling File exist Function
		WScript.Echo "File Not Exist in "& excelFilePath		
		fReadExcelFile = "File Not Exist in "& excelFilePath
		Exit Function
	End If
	
	Set objXls = CreateObject("Excel.Application")
    Set objWBook = objXls.Workbooks.Open(excelFilePath)
    Set objWSheet = objWBook.Worksheets(excelSheetName)
				
	Set END_ROW = objWSheet.Range("B1:B20000").Find("END")	
	
	If Not END_ROW Is Nothing Then
  		strValue =  END_ROW.Row  
	End If
	
	Set PARAM_Cell_COLUMN = objWSheet.Range("A1:AZ1").Find("PARAM0")
	
	If Not PARAM_Cell_COLUMN Is Nothing Then
		Wscript.Echo "PARAM01 Doesn't found!"
	End If
        
	Set Found_Empty_Row = objWSheet.Range("B1:B"&strValue).Find("")
	
	If Not Found_Empty_Row Is Nothing Then
	
		
       	For j = PARAM_Cell_COLUMN.Column  To 500		

        If Trim(objWSheet.Cells(Found_Empty_Row.Row, j))<>"" then
		  
			If Trim(objWSheet.Cells(Found_Empty_Row.Row + 1, j))="" Then			
				Conc_cnt="Empty"
			Else				
				Conc_cnt=Trim(objWSheet.Cells(Found_Empty_Row.Row + 1, j))
				
			End If	
					   	
	        Data_Var = Data_Var & Trim(objWSheet.Cells(Found_Empty_Row.Row, j)) & "," & Conc_cnt & ","
			
		Else 
			Exit For
        	End if        
		Next        
        Data_Var = Left(Data_Var, Len(Data_Var) - 1)	    
		SNO=Data_Var
	Else
		SNO = "none"		
	End If
	
	objWBook.Save	
	objWBook.Close
	objXls.Quit
	
	Set objWSheet = Nothing
	Set objXls = Nothing
	Set Found_Empty_Row = Nothing
	set Copyrange = Nothing

fReadExcelFile = SNO
End Function

'############################################################

' Function name: fReadLibreFile
' Description:   Reading LibreOffice File 
' Parameters:  Excel File Path and Excel Sheet Name
'#############################################################

Function fReadLibreFile(excelFilePath,excelSheetName)

	Dim OpenPar(2)                                 'Variables for Handing LibreOffice file Object
	Dim wb 
	
	If FileExist(excelFilePath) = False Then ' Calling File exist Function
		WScript.Echo "File Not Exist in "& excelFilePath
		fReadLibreFile = "File Not Exist in "& excelFilePath
		Exit Function
	End If

	Set oSM = CreateObject("com.sun.star.ServiceManager")
	Set oDesk = oSM.createInstance("com.sun.star.frame.Desktop")
		
	Set OpenPar(0) = MakePropertyLibre("ReadOnly", True)  ' Setting File Read only using MakePropertyLibre Function
	Set OpenPar(1) = MakePropertyLibre("Hidden", True)	  ' Setting File Hidden using MakePropertyLibre Function
	Set wb = oDesk.loadComponentFromURL("file:///" & excelFilePath, "_blank", 0, OpenPar)
	Set objWSheet = wb.getSheets().getByName(excelSheetName)
			
	For i_End = 0 to 1000   										' For loop for getting the END Row Number
		If objWSheet.getCellByPosition(1, i_End).String = "END" Then
			END_ROW = i_End
			Exit For
		End If
	Next
	
	For i_EmptyRow = 0 to END_ROW									' For loop for getting the Empty Row Number
		If objWSheet.getCellByPosition(1, i_EmptyRow).String = "" Then			
			Found_Empty_Row = i_EmptyRow 
			Exit For
		End If
	Next
		
	For i_LastCol=3  to 1000										' For loop for getting the Last Column of Calender
		If objWSheet.getCellByPosition(i_LastCol,Found_Empty_Row).String = "" Then
			Found_Last_Column = i_LastCol
			Exit For
		End If
	Next

	For i_par = 2 to 1000											' For loop for getting the Column Number of PARAM0
			If objWSheet.getCellByPosition(i_par,0).String ="PARAM0" Then			
				PARAM_Cell_COLUMN = i_par
				Exit For
			End If
	Next
	
	If PARAM_Cell_COLUMN = "" Then
		Wscript.Echo "PARAM0 Doesn't found!"
		msgbox "PARAM0 Doesn't found!"
	End If
	
	if Found_Empty_Row <> "" then
	  
		For  i_conc = PARAM_Cell_COLUMN to Found_Last_Column
		
			If Trim(objWSheet.getCellByPosition(i_conc,Found_Empty_Row).String) <> "" then
			
				If Trim(objWSheet.getCellByPosition(i_conc,Found_Empty_Row+1).String) ="" then
					Conc_cnt="Empty"
				Else
					Conc_cnt=Trim(objWSheet.getCellByPosition(i_conc,Found_Empty_Row+1).String)
				End if
				
			Data_Var = Data_Var & Trim(objWSheet.getCellByPosition(i_conc,Found_Empty_Row).String) & "," & Conc_cnt & ","
			Else
			Exit For
			End if
	    Next
		Data_Var = Left(Data_Var, Len(Data_Var) - 1)
		SNO=Data_Var
	Else
		SNO = "none"
	End If
	
	oDesk.terminate                                'Terminating the LibreOffice Object
	Set objWSheet = Nothing
	Set oSM = Nothing
	Set Found_Empty_Row = Nothing
	set Copyrange = Nothing
	
fReadLibreFile = SNO

End Function

If Right(excelFilePath,4) = ".ods" Then
	ReadData = fReadLibreFile(excelFilePath,excelSheetName)
elseif  Right(excelFilePath,5) = ".xlsx" Then
	ReadData = fReadExcelFile(excelFilePath,excelSheetName)	
else
	
	Wscript.Echo "Invalid File format!"
End If
Wscript.echo "~~~GINGER_RC_START~~~"
StdOut.Write("Outputvalue =")
Wscript.StdOut.Write(ReadData)
Wscript.echo "~~~GINGER_RC_END~~~"