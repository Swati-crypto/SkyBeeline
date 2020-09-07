'GINGER_Description Renamefile
'GINGER_$excelFilePath
'GINGER_$excelSheetName
'GINGER_$PARAM_NAME
'GINGER_$PARAM_VALUE
'GINGER_$GROUP_NAME

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
Dim PARAM_VALUE
Dim GROUP_NAME
Dim intRow : intRow = 1

excelFilePath = WScript.Arguments(0)
excelSheetName = WScript.Arguments(1)
PARAM_NAME = WScript.Arguments(2)
PARAM_VALUE = WScript.Arguments(3)
GROUP_NAME = trim(WScript.Arguments(4))
'msgbox GROUP_NAME
'msgbox len(GROUP_NAME)


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

' Function name: fReadDatafromDataFile

' Description:   

' Parameters:    None

' Return value:  Success - True , Fail - False                           



'#############################################################

 Function fReadDatafromDataFile(excelFilePath)

	If FileExist(excelFilePath) = False Then
		WScript.Echo "File does not exist at "& excelFilePath		
		fReadDatafromDataFile = "File does not exist at "& excelFilePath
		Exit Function
	End If

	
	Set objXls = CreateObject("Excel.Application")
	Set objWBook = objXls.Workbooks.Open(excelFilePath)
	Set objWSheet = objWBook.Worksheets(excelSheetName)

	If excelSheetName ="KEEP_REFER" then	
		intLength = Len(PARAM_NAME)
		PARAM_NAME = mid(PARAM_NAME, 2, intLength-1)
		cnt_rows = objWBook.WorkSheets("KEEP_REFER").usedRange.rows.count
		'msgbox cnt_rows
  		Set FoundCell_VarName= objWBook.WorkSheets("KEEP_REFER").Range("A1:A65000").FIND(PARAM_NAME)	

			If Not FoundCell_VarName Is Nothing Then					
				'Msgbox PARAM_NAME
				objWBook.WorkSheets("KEEP_REFER").Range("B" & FoundCell_VarName.Row).Value = PARAM_VALUE
			Else	
				objWBook.WorkSheets("KEEP_REFER").Range("A" & cnt_rows+1).Value = PARAM_NAME
				objWBook.WorkSheets("KEEP_REFER").Range("B" & cnt_rows+1).Value = PARAM_VALUE
'				Conc_Str="Parameter_Not_Found"		
'				MsgBox "Parameter_Not_Found :" &  PARAM_NAME	
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
		'msgbox strValue
	

	End If

        Set FoundCell_Empty = objWSheet.Range("B1:B"&strValue).Find("")
		'msgbox FoundCell_Empty.Row
On error resume Next

	
	Set FoundCell = objWSheet.Range("A" & FoundCell_Empty.Row & ":BZ" & FoundCell_Empty.Row).Find(PARAM_NAME,,,1)	


	If Not FoundCell_Empty Is Nothing Then

		'SNO = objWSheet.Cells(FoundCell.Row,1)
		If PARAM_NAME="SKIP" AND PARAM_VALUE<>"N" Then
		'msgBox FoundCell_Empty.Row
					
		Set FoundCell_colNumber_flow_Control = objWSheet.Range("A1:ZZ"&strValue).Find("FLOW_CONTROL")		
			msgBox FoundCell_colNumber_flow_Control
			
			int_Flow_Control_Col = FoundCell_colNumber_flow_Control.Column
			msgBox int_Flow_Control_Col
		
			FLOW_CONTROL =  objWSheet.Cells(FoundCell_Empty.Row+1,int_Flow_Control_Col).Value
			msgBox FLOW_CONTROL
			
			objWSheet.Cells(FoundCell_Empty.Row,2).Value="X" 		
			objWSheet.Cells(FoundCell_Empty.Row+1,2).Value=PARAM_VALUE
			
			If PARAM_VALUE ="P" Then
			msgBox "XYZ"
			Set FoundCell_FirstID = objWSheet.Range("A1:A" & strValue).Find("1") 
			msgBox FoundCell_FirstID
			Set FoundCell_colNumber = objWSheet.Range("A1:ZZ"&strValue).Find("GROUP_NAME")	''
			msgBox FoundCell_colNumber
			
			msgBox FoundCell_FirstID
			msgBox strValue
			
			'intRow = intRow + 1
			'msgBox intRow
			
			int_Group_Name_Col = FoundCell_colNumber.Column
			msgBox int_Group_Name_Col''
			intRow = FoundCell_FirstID.Row + 1
			msgBox intRow
			
			For intRow = FoundCell_FirstID.Row + 1 to strValue-1
			msgBox "forloop"
			'Set FoundCell_FirstID = objWSheet.Range("A1:A"&strValue).Find("1")
			'Set FoundCell_colNumber = objWSheet.Range("A1:ZZ"&strValue).Find("GROUP_NAME")			
			'msgbox FoundCell_colNumber.Column
			
			'msgbox intRow
			'msgbox objWSheet.Cells(intRow,2)
			'msgbox objWSheet.Cells(intRow,6)
			'msgbox GROUP_NAME
			'msgbox trim(objWSheet.Cells(intRow,intCol))
			'Set a= trim(objWSheet.Cells(intRow,intCol))
			'msgBox a
			msgbox trim(objWSheet.Cells(intRow,int_Group_Name_Col))
			msgBox trim(objWSheet.Cells(intRow,2))
			msgBox trim(objWSheet.Cells(intRow,int_Flow_Control_Col))
				If trim(objWSheet.Cells(intRow,2)) = "" AND trim(objWSheet.Cells(intRow,int_Group_Name_Col)) = GROUP_NAME AND trim(objWSheet.Cells(intRow,int_Flow_Control_Col)) = FLOW_CONTROL Then
				
				msgBox "write"
					objWSheet.Cells(intRow-1,2).Value = "X"
					objWSheet.Cells(intRow,2).Value = "N"
					'Exit For
				End If				
				intRow = intRow + 1
				msgBox "increament intRow"
				If intRow >= strValue Then								
					Exit For
				End If
				msgBox "Next"
				
			Next
			End If
			
		
		ElseIf PARAM_NAME="SKIP" AND PARAM_VALUE="N" Then
			'msgbox strValue
			'msgbox objWSheet.Cells(3,2).Value
			'msgbox objWSheet.Cells(11,2).Value
			Set FoundCell_FirstID = objWSheet.Range("A1:A"&strValue).Find("1")
			Set FoundCell_colNumber = objWSheet.Range("A1:ZZ"&strValue).Find("GROUP_NAME")			
			'msgbox FoundCell_colNumber.Column
			intCol = FoundCell_colNumber.Column
			'msgBox intCol
			intRow = FoundCell_FirstID.Row + 1
				'msgBox intRow
			
			'msgbox strValue
			For intRow = FoundCell_FirstID.Row + 1 to strValue-1
			
			'msgbox intRow
			'msgbox objWSheet.Cells(intRow,2)
			'msgbox objWSheet.Cells(intRow,6)
			'msgbox GROUP_NAME
			'msgbox trim(objWSheet.Cells(intRow,intCol))
			'Set a= trim(objWSheet.Cells(intRow,intCol))
			msgBox trim(objWSheet.Cells(intRow,intCol))
			msgBox trim(objWSheet.Cells(intRow,2))
				If trim(objWSheet.Cells(intRow,2)) = "" AND trim(objWSheet.Cells(intRow,intCol)) = GROUP_NAME Then
				msgBox "Write"
					objWSheet.Cells(intRow-1,2).Value = "X"
					objWSheet.Cells(intRow,2).Value = "N"
					'Exit For
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

fReadDatafromDataFile = SNO
 End Function

'*************************************************************************************************************************************************

Call FileExist(excelFilePath)

 ReadData = fReadDatafromDataFile(excelFilePath)



' output results to Ginger
Wscript.echo "~~~GINGER_RC_START~~~" 
WScript.Echo ReadData 
'Wscript.echo strVariable & "=" + ReadData 
Wscript.echo "~~~GINGER_RC_END~~~"














