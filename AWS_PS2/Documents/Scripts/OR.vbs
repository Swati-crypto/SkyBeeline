'GINGER_$excelORFilePath
'GINGER_$excelSheetName
'GINGER_$testCaseName"
'GINGER_$StorageORFilePath

if WScript.Arguments.Count = 0 then
    WScript.Echo "Missing parameters"        
end if

excelORFilePath = WScript.Arguments(0)
excelSheetName = WScript.Arguments(1)
testCaseName = WScript.Arguments(2)
StorageORFilePath = WScript.Arguments(3)
'Get OR FIle from OR File Path
excelFilePath = excelORFilePath
arrGetExcelFile = Split(excelORFilePath,"\")
getExcelFile = arrGetExcelFile(Ubound(arrGetExcelFile))
strDestinationLocation = left(excelFilePath,(len(excelFilePath)-len(getExcelFile)))

strORSourcexlsLocation = StorageORFilePath &"\"& getExcelFile

Set objFS=CreateObject("Scripting.FilesystemObject") 

'Verify OR Excel file exist in Storage folder, if it not exist then stop execution
If objFS.FileExists(strORStoragepath &"\"& getExcelFile) = False Then        
    xlsORFlag=0
    ScriptStatus = getExcelFile & "Not Exist in Storage Folder" & strORStoragepath     
    ErrorCode = "OR Excel File Not found in Storage !!" 
End If

'Check If OR File Is Already placed at correct location 
If (objFS.FileExists(excelFilePath)) Then 		
 		objFS.DeleteFile(excelFilePath)
End If

 'Copy the OR file from source to destination
    If objFS.FileExists(excelORFilePath &"\"& strORxlsFileName) = False Then        
        objFs.CopyFile strORSourcexlsLocation, strDestinationLocation,False
    Else
        ErrorCode = "File not copied"
    End If



Set objXls = CreateObject("Excel.Application")
Set objWBook = objXls.Workbooks.Open(excelFilePath)
Set objWSheet = objWBook.Worksheets(excelSheetName)
	
'Remove OR.xlsx file in Execution folder. created Function for "verifyLogFile"
excelORFilePath = strDestinationLocation
verifyLogFile excelORFilePath

str=""
Dim rowNumber
on error resume next
	' Set FoundCell = objWSheet.Range("B1:B50000" & strValue).Find(testCaseName)
      Set FoundCell= objWSheet.Range("B1:B50000").Find(testCaseName,,,1)
        
      rowNumber = FoundCell.Row
	'	rowNumber = objWSheet.Range("B1:B50000").Find(testCaseName,,,1).Row
        rowCnt = objWSheet.Usedrange.columns.Count
        
        For r = 3 To rowCnt
        
        	If IsEmpty(objWSheet.Cells(rowNumber,r).value) = False Then
        		
        		Concatinate = Concatinate & "," & objWSheet.Cells(rowNumber,r).value
        			
        	End If
        	
        Next
        
        TC_concatinate =  Right(Concatinate,len(Concatinate)-1)
      kr_row_Cnt = objWBook.Worksheets("KEEP_REFER").usedrange.Rows.count
   		kr_col_Cnt = objWBook.Worksheets("KEEP_REFER").usedrange.columns.count
   		
   		For c = 1 To kr_col_Cnt
   		
   			If objWBook.Worksheets("KEEP_REFER").cells(1,c) = "LOCATE_NAME"  Then
   				lName = c
   			End If
   			
   			If objWBook.Worksheets("KEEP_REFER").cells(1,c) = "LOCATE_VALUE"  Then
   				lValue = c
   			End If
   			
   		Next
   		
   		Set DictionaryObject = CreateObject("Scripting.Dictionary")
   		
   		For i=1 to kr_row_Cnt
   			locateName = objWBook.Worksheets("KEEP_REFER").cells(i,lName).value
   			locateValue = objWBook.Worksheets("KEEP_REFER").cells(i,lValue).value
   			if DictionaryObject.Exists(locateName) =False Then
   				DictionaryObject.Add locateName,locateValue
   			End If
   		
   		Next
   		
   		arr_ORValue=Split(TC_concatinate,",")
        
        For Iterator = 0 To Ubound(arr_ORValue)
            arr_locaterName = arr_ORValue(Iterator)
        	
					If DictionaryObject.Exists(arr_locaterName) = True Then
						If DictionaryObject.Item(arr_locaterName) <> "" Then
							con_str = DictionaryObject.Item(arr_locaterName)
						Else
							con_str ="Locater_Value_is_Empty"
						End If
						
					Else
						con_str ="Locater_Name_Not_Found"
					End If
		        	    
		        	str=str & trim(arr_locaterName) &"/$/" & con_str &"/$/" 
		        	
        			writeDataLog = arr_locaterName &"," & con_str
        			
        			'Write Data into Log file
        			WriteDateinTextFile excelORFilePath,writeDataLog,testCaseName
       Next
        
        strORReadData = Replace(str,"=","|")
        
              
'        On error goto 0
        
  DictionaryObject.RemoveAll 
	objWBook.Save
	
	objWBook.Close

	objXls.Quit
	
	 If (objFS.FileExists(excelFilePath)) Then 	
     	objFS.DeleteFile(excelFilePath)
	End If
	Set objFS = Nothing
	Set objWSheet = Nothing
	Set objXls = Nothing
	Set Key_No = Nothing
	set Copyrange = Nothing
	
'===========================================================================================================================================================
'Function Name: Create Excel file
'Creation Date : 4-Apr-2017
'===========================================================================================================================================================
Function WriteDateinTextFile(strTextFilePath,WriteData,testCaseName)
	
		Set fso = CreateObject("Scripting.FileSystemObject")
		If Right(strTextFilePath,1) = "\" Then
			textFilePath = strTextFilePath & "OR.log"
		Else
			textFilePath = strTextFilePath &"\OR.log"
		End If
	If fso.FileExists(textFilePath) = False Then
		Set fs = fso.CreateTextFile(textFilePath,8)
	Else
		Set fs = fso.OpenTextFile(textFilePath,8,True)
	End If
	
	arr_writetoORlog = Split(writeDataLog,",")
	locater_Name = arr_writetoORlog(0)
	locater_Des = arr_writetoORlog(1)
	'Set fs = fso.OpenTextFile(textFilePath)

	
		If locater_Des = "Locater_Value_is_Empty" Then
			fs.WriteLine "***********************************************************************"
			fs.WriteLine "["&testCaseName &"]->["& locater_Name & "] => Locater value Not Found in Keep Refer Sheet"
			fs.WriteLine "***********************************************************************"
		End If
		
		If locater_Des = "Locater_Name_Not_Found" Then
			fs.WriteLine "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
			fs.WriteLine "["&testCaseName &"]->["& locater_Name & "] => Locater Name Not exist in Keep Refer Sheet"
			fs.WriteLine "%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%%"
		End If

Set fs = Nothing
Set fso=Nothing
	
		
End Function

'===========================================================================================================================================================
'Function Name: Verify Log file and Delete file
'Creation Date : 4-Apr-2017
'===========================================================================================================================================================
Function verifyLogFile(strTextFilePath)
	
	Set fso = CreateObject("Scripting.FileSystemObject")
		If Right(strTextFilePath,1) = "\" Then
			textFilePath = strTextFilePath & "OR.log"
		Else
			textFilePath = strTextFilePath &"\OR.log"
		End If
	
	If fso.FileExists(textFilePath) = True Then
		fso.DeleteFile(textFilePath)
	End If

Set fs = Nothing
Set fso=Nothing
	
		
End Function



' output results to Ginger
Wscript.echo "~~~GINGER_RC_START~~~" 
WScript.Echo strORReadData 
'Wscript.echo strVariable & "=" + ReadData 
Wscript.echo "~~~GINGER_RC_END~~~"
