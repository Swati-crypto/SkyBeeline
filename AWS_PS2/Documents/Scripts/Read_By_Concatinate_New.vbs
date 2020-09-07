'GINGER_$excelFilePath
'GINGER_$excelSheetName

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
Dim str
Dim Conc_Str
Dim Temp_VAr
Dim nLastCol

excelFilePath = WScript.Arguments(0)
excelSheetName = WScript.Arguments(1)

'############################################################
' Modify By: Naresh
' Description:   New object creation for "StdOut.Write" 
' Modification date:  01 August                          
'#############################################################

SET FS = CreateObject("Scripting.FileSystemObject")
SET StdOut = FS.GetStandardStream(1)

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

	
	str=""		
	Set FoundCell = objWSheet.Range("B1:B20000").Find("END")	
	
	If Not FoundCell Is Nothing Then
  		strValue =  FoundCell.Row          
	End If

     Set FoundCell_Empty = objWSheet.Range("B1:B" & strValue).Find("")
	
On Error resume Next	 

	Set FoundCell = objWSheet.Range("A" & FoundCell_Empty.Row & ":BZ" & FoundCell_Empty.Row).Find(PARAM_NAME, , , 1)
	

	If Not FoundCell Is Nothing Then

	If Not FoundCell_Empty Is Nothing Then		
			       
		If PARAM_NAME="SKIP" Then
			
			objWSheet.Cells(FoundCell_Empty.Row,2).Value="X" 		
			objWSheet.Cells(FoundCell_Empty.Row+1,2).Value=PARAM_VALUE 		
		Else
						
			EndXl_Col=objWSheet.Cells(FoundCell_Empty.Row,500).End("-4159").Column
			
			'Msgbox EndXl_Col
						
			For i=3 to EndXl_Col
					
				If Instr(objWSheet.Cells(FoundCell_Empty.Row,i),"SCREEN_DATA_APP") > 0 then

					i=i+1
				End if
					
				If objWSheet.Cells(FoundCell_Empty.Row + 1,i)="" then

					Conc_Str="Value is Blank"
				Else
					Conc_Str=objWSheet.Cells(FoundCell_Empty.Row + 1,i)
					
				   If  Mid(Conc_Str,1,1)="@" then
					
					Set FoundCell_VarName= objWBook.WorkSheets("KEEP_REFER").Range("A1:A65000").FIND(Mid(objWSheet.Cells(FoundCell_Empty.Row+1,i),2,Len(objWSheet.Cells(FoundCell_Empty.Row+1,i))))	

					    If FoundCell_VarName Is Nothing Then
						
						EndXl_Row=objWBook.WorkSheets("KEEP_REFER").Range("A65000").End("-4162").Row
	
						objWBook.WorkSheets("KEEP_REFER").Range("A" & EndXl_Row+1)= Mid(objWSheet.Cells(FoundCell_Empty.Row+1,i),2,Len(objWSheet.Cells(FoundCell_Empty.Row+1,i)))
					    End If	

				  Elseif Mid(Conc_Str,1,1)="&" then	
						
					    Set FoundCell_VarName= objWBook.WorkSheets("KEEP_REFER").Range("A1:A65000").FIND(Mid(objWSheet.Cells(FoundCell_Empty.Row+1,i),2,Len(objWSheet.Cells(FoundCell_Empty.Row+1,i))))	

					    If Not FoundCell_VarName Is Nothing Then
						
						Conc_Str= objWBook.WorkSheets("KEEP_REFER").Range("B" & FoundCell_VarName.Row)
					    Else
						
						Conc_Str="Parameter_Not_Found"		
						MsgBox "Parameter_Not_Found :" &  Mid(objWSheet.Cells(FoundCell_Empty.Row+1,i),2,Len(objWSheet.Cells(FoundCell_Empty.Row+1,i)))		
					    End If	
				   End if 

				End if

				Str= Str + objWSheet.Cells(FoundCell_Empty.Row,i) + "/$/" +  trim(Conc_Str) + "/$/" 
			'Ne:

			Next
		End If 	  					   	
	  
	Else
		str = "none"

	End If
	End If
	
	objWBook.Save
	objWBook.Close
	objXls.Quit
	
	Set objWSheet = Nothing
	Set objXls = Nothing
	Set FoundCell = Nothing
	set Copyrange = Nothing

	fReadDatafromDataFile = str
	
 End Function
 
'############################################################
' Function name: fgetSQLqueryfromCommonfile
' Description:   
' Parameters:    None
' Return value:  Success - True , Fail - False                           
'#############################################################

	Function fgetSQLqueryfromCommonfile(excelFilePath,strText)
	
		arr_excelfilepath = split(excelFilePath,"\")
		strCommonXlspath = left(excelFilePath,len(excelFilePath)-len(arr_excelfilepath(ubound(arr_excelfilepath))))& "COMMON.xls"
	
	
	 Set objXls = CreateObject("Excel.Application")
   Set objWBook = objXls.Workbooks.Open(strCommonXlspath)
   Set objWSheet = objWBook.Worksheets("DB_SQL")

		colCnt = objWSheet.usedRange.columns.count
	
		
		For c = 1 To colCnt
			
			If objWSheet.cells(1,c).value = "PK" Then
				col_pk = c
			End If
			
			If objWSheet.cells(1,c).value = "SQL_1" Then
				col_sql1 = c
			End If
			
		Next
			
		On Error resume Next	 

	Set GlobalDictionary = CreateObject("Scripting.Dictionary")			'Adding all the variables an values are stored in Dictionary objects

		strcomText = Split(strText,"/$/")
		arr_strcomText = ubound(strcomText)
		
		For i = 0 To arr_strcomText-1
			incArrayValue = i+1
			
			If GlobalDictionary.Exists(strcomText(i)) = False Then
				GlobalDictionary.Add strcomText(i),strcomText(incArrayValue)
			End If
			
			i=incArrayValue
		Next
		
		incPK = 1
		For i = 0 To arr_strcomText-1					'Verify PK values and updating the values
			incArrayValue = i+1
			
			If strcomText(i) = "PK" Then
				pkFlag = 1
				PK_VALUE = GlobalDictionary(strcomText(i))
			Else
				If strcomText(i)="PK_"&incPK Then
					pkFlag = 1	
					PK_VALUE = GlobalDictionary(strcomText(i))
					incPK = incPK+1					
				End If	
				
			End If		
			
			If pkFlag = 1 Then
				
				Set FoundCell = objWSheet.Range("B1:B10000").Find(PK_VALUE, , , 1)	'Finding the PK value in Common xls file
				foundcellRow = FoundCell.row
				
				SQL_modify = objWSheet.cells(foundcellRow,col_sql1).value
				col_sql2 = col_sql1+1
				col_sql3 = col_sql2+1
				col_sql4 = col_sql3+1
				
				If objWSheet.cells(foundcellRow,col_sql2).value <> "" Then								'Concatinate SQL_2,SQL_3 & SQL_4 columns of data
					SQL_modify = SQL_modify & objWSheet.cells(foundcellRow,col_sql2).value							
				End If
				
				If objWSheet.cells(foundcellRow,col_sql3).value <> "" Then
					SQL_modify = SQL_modify & objWSheet.cells(foundcellRow,col_sql3).value							
				End If
				
					If objWSheet.cells(foundcellRow,col_sql4).value <> "" Then
					SQL_modify = SQL_modify & objWSheet.cells(foundcellRow,col_sql4).value							
				End If
				
					'-------------------------------------------------------------------------------------------
						range1 = 1
						range2 = 1
						
						For j = 1 To 20
								
							range1 = instr(range1+1 , SQL_modify , "<")
							
							If range1 = 0 Then
								Exit For
							End If
						
							range2 = instr(range2+1 , SQL_modify , ">")
							
							If range2 = 0 Then
								Exit For
							End If
							
							diffRange = range2 - range1
						
							Variable = mid(SQL_modify, range1+1 , diffRange-1)
							
							var = split(trim(Variable)," ")
											
							n = ubound(var)
							
							If n = 0 Then
								
								Temp = 	GlobalDictionary.Item(Variable)	
								'Temp = objWSheet.Cells(rownum,2).value
								
								SQL_modify =  replace(SQL_modify , "<"& Variable &">" , Temp) 
								
							End If			
							
						Next 
					
					'-------------------------------------------------------------------------------------------
				
			End If
				strText = Replace(strText,PK_VALUE,SQL_modify,1,1,0)
			
			i=incArrayValue
		Next	
		
	objWBook.Save
	objWBook.Close
	objXls.Quit
	
	Set objWSheet = Nothing
	Set objXls = Nothing
	Set FoundCell = Nothing
	set Copyrange = Nothing
	
	fgetSQLqueryfromCommonfile = strText
	
	 End Function
	
'========================================================================================================================================
'############################################################
' Function name: fgetUnixCommandfromCommonfile
' Description:   
' Parameters:    None
' Return value:  Success - True , Fail - False                           
'#############################################################
Function fgetUnixCommandfromCommonfile(excelFilePath,strText)
	
			arr_excelfilepath = split(excelFilePath,"\")
			strCommonXlspath = left(excelFilePath,len(excelFilePath)-len(arr_excelfilepath(ubound(arr_excelfilepath))))& "COMMON.xls"
			
			
			Set objXls = CreateObject("Excel.Application")
		   	Set objWBook = objXls.Workbooks.Open(strCommonXlspath)
		   	Set objWSheet = objWBook.Worksheets("JOBS")
		
				colCnt = objWSheet.usedRange.columns.count
			
				
				For c = 1 To colCnt
					
					If objWSheet.cells(1,c).value = "PROCESS_NAME" Then
						col_PN = c
					End If
					
					If objWSheet.cells(1,c).value = "JOB_NAME" Then
						col_JN = c
					End If
					
					If objWSheet.cells(1,c).value = "JOB_REC" Then
						col_JR = c
					End If
					
					If objWSheet.cells(1,c).value = "COMMAND" Then
						col_CMD = c
					End If
					
					If objWSheet.cells(1,c).value = "EXPECTED_STRING" Then
						col_EXPSTR = c
					End If
					
					If objWSheet.cells(1,c).value = "PRIOR_WAIT_TIME" Then
						col_PR_WAIT_TIME = c
					End If
					
					If objWSheet.cells(1,c).value = "EXECUTION_TIME" Then
						col_EXE_TIME = c
					End If
					
					If objWSheet.cells(1,c).value = "JOB_INTER_CHANGE" Then
						col_JOB_INTER = c
					End If				
				Next
					
				On Error resume Next	 
		
			Set GlobalDictionary = CreateObject("Scripting.Dictionary")			'Adding all the variables an values are stored in Dictionary objects
		
				strcomText = Split(strText,"/$/")
				arr_strcomText = ubound(strcomText)
				
				For i = 0 To arr_strcomText-1
					incArrayValue = i+1
					
					If GlobalDictionary.Exists(strcomText(i)) = False Then
						GlobalDictionary.Add strcomText(i),strcomText(incArrayValue)
					End If
					
					i=incArrayValue
				Next
				
				incPN = 1
				For i = 0 To arr_strcomText-1					'Verify PK values and updating the values
					incArrayValue = i+1
					
					If strcomText(i) = "PROCESS_NAME" Then
						pnFlag = 1
						PN_VALUE = GlobalDictionary(strcomText(i))
					Else
						If strcomText(i)="PROCESS_NAME_"&incPN Then
							pnFlag = 1
							PN_VALUE = GlobalDictionary(strcomText(i))
							PN_inc = incPN
							incPN = incPN+1					
						End If	
						
					End If		
					
					If pnFlag = 1 Then
						
						Set FoundCell = objWSheet.Range("A1:A10000").Find(PN_VALUE, , , 1)	'Finding the PK value in Common xls file
						foundcellRow = FoundCell.row
						
				
						strJob_Name = objWSheet.cells(foundcellRow,col_JN).value
						
							If len(strJob_Name) <> 0 Then
								process_value = strJob_Name
							End If
							strJob_Rec = objWSheet.cells(foundcellRow,col_JR).value
							
							strJob_InterChange = objWSheet.cells(foundcellRow,col_JOB_INTER).value							
							
							If Ucase(strJob_InterChange) <> "Y" Then
								If len(strJob_Rec) <> 0 Then
									process_value = strJob_Name &" "& strJob_Rec
								End If
							Else
								If len(strJob_Rec) <> 0 Then
									process_value = strJob_Rec &" "& strJob_Name
								End If						
							
							End If
							
						strCMD = objWSheet.cells(foundcellRow,col_CMD).value
									
						strExep_String = objWSheet.cells(foundcellRow,col_EXPSTR).value
						
							If len(strExep_String) <> 0 Then
								process_value = process_value& "/$/EXPECTED_STRING_"&PN_inc &"/$/" & strExep_String
							End If
							
						strPirior_Wait_time = objWSheet.cells(foundcellRow,col_PR_WAIT_TIME).value
							If len(strPirior_Wait_time) <> 0 Then
								process_value = process_value& "/$/PRIOR_WAIT_TIME_"&PN_inc &"/$/" & strPirior_Wait_time
							End If
							
						strExe_Wait_time = objWSheet.cells(foundcellRow,col_EXE_TIME).value
							If len(strExe_Wait_time) <> 0 Then
								process_value = process_value& "/$/EXECUTION_TIME_"&PN_inc &"/$/" & strExe_Wait_time
							End If
							
							
						
							'-------------------------------------------------------------------------------------------
								range1 = 1
								range2 = 1
								
								For j = 1 To 20
										
									range1 = instr(range1+1 , strCMD , "<")
									
									If range1 = 0 Then
										Exit For
									End If
								
									range2 = instr(range2+1 , strCMD , ">")
									
									If range2 = 0 Then
										Exit For
									End If
									
									diffRange = range2 - range1
								
									Variable = mid(strCMD, range1+1 , diffRange-1)
									
									var = split(trim(Variable)," ")
													
									n = ubound(var)
									
									If n = 0 Then
										
										Temp = 	GlobalDictionary.Item(Variable)	
										'Temp = objWSheet.Cells(rownum,2).value
										
										strCMD =  replace(strCMD , "<"& Variable &">" , Temp) 
										
									End If			
									
								Next 
							
							'-------------------------------------------------------------------------------------------
						
					End If
						If len(strCMD) <> 0 Then
							process_value  =strCMD & process_value
						End If
						strText = Replace(strText,PN_VALUE,process_value,1,1,0)
						print strText
						process_value=""
						strCMD = ""
					
					i=incArrayValue
				Next	
				
			objWBook.Save
			objWBook.Close
			objXls.Quit
			
			Set objWSheet = Nothing
			Set objXls = Nothing
			Set FoundCell = Nothing
			set Copyrange = Nothing
			
			fgetUnixCommandfromCommonfile = strText
	
	 End Function

'*************************************************************************************************************************************************

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

' Function name: fReadLibreFile
' Description:   Reading Libre Office File 
' Parameters:  Excel File Path and Excel Sheet Name
'#############################################################

Function fReadLibreFile(excelFilePath,excelSheetName)

	Dim OpenPar(2)
	Dim wb
	
	If FileExist(excelFilePath) = False Then ' Calling File exist Function
		WScript.Echo "File Not Exist in "& excelFilePath		
		fReadDatafromDataFile = "File Not Exist in "& excelFilePath
		Exit Function
	End If	

	Set oSM = CreateObject("com.sun.star.ServiceManager")
	Set oDesk = oSM.createInstance("com.sun.star.frame.Desktop")
		
	Set OpenPar(0) = MakePropertyLibre("ReadOnly", True)  ' Setting File Read only using MakePropertyLibre Function
	Set OpenPar(1) = MakePropertyLibre("Hidden", True)	  ' Setting File Hidden using MakePropertyLibre Function
	Set wb = oDesk.loadComponentFromURL("file:///" & excelFilePath, "_blank", 0, OpenPar) 
	Set oSheet = wb.getSheets().getByName(excelSheetName)
	Set oSheet2 = wb.getSheets().getByName("KEEP_REFER")
	
	For i_End = 0 to 1000
		If oSheet.getCellByPosition(1, i_End).String = "END" Then
			FoundCell = i_End 
			Exit For
		End If
	Next	
		
	bFlagEmptyCellFound  = false
	
	For i_EmptyRow = 0 to FoundCell
		If oSheet.getCellByPosition(1, i_EmptyRow).String = "" Then
			bFlagEmptyCellFound  = true
			FoundCellEmpty = i_EmptyRow 
			Exit For
		End If
	Next
	
	For i_LastCol=3  to 1000
		If oSheet.getCellByPosition(i_LastCol,FoundCellEmpty).String = "" Then
			bFlagLstColumn  = true
			FLstColumn = i_LastCol 
			Exit For
		End If
	Next	
	
	If bFlagEmptyCellFound Then
		For i_par = 0 to 15
			If oSheet.getCellByPosition(i_par,0).String ="TEST_NAME" Then
			
				PARAM_Col_num = i_par
			
				Exit For
			End If
	Next	
			
	Else
		
	End If 
		
	For  i_conc = PARAM_Col_num to FLstColumn-1

		If oSheet.getCellByPosition(i_conc,FoundCellEmpty+1).String ="" then		
			
		Else 				
				If left(oSheet.getCellByPosition(i_conc,FoundCellEmpty+1).String,1)="&" Then
					'mSGBOX oSheet.getCellByPosition(i_conc,FoundCellEmpty+1).String
					
				
					For i_EmptyRow1 = 1 to 500
						
						If "&" & oSheet2.getCellByPosition(0,i_EmptyRow1).String = oSheet.getCellByPosition(i_conc,FoundCellEmpty+1).String Then
							PARAM_VALUE= PARAM_VALUE & "/$/" & oSheet.getCellByPosition(i_conc,FoundCellEmpty).String & "/$/" &  oSheet2.getCellByPosition(1,i_EmptyRow1).String
							
							exit for
						End If	
					Next
						
				
				Else
				
					PARAM_VALUE= PARAM_VALUE & "/$/" & oSheet.getCellByPosition(i_conc,FoundCellEmpty).String & "/$/" &  oSheet.getCellByPosition(i_conc,FoundCellEmpty+1).String
					
				End If
		End if 	
	Next
			
			
		
	
        PARAM_VALUE=PARAM_VALUE & "/$/"
		
	oDesk.terminate
	Set oSM = Nothing
	Set oSheet = Nothing

fReadLibreFile = PARAM_VALUE 

End Function

'############################################################

' Modify By: Naresh
' Description:  Method to get value out from VBS  -->  StdOut.Write("Outputvalue =")
' Modification date:  01 August                          

'#############################################################

If Right(excelFilePath,4) = ".ods" Then
	
	ReadData = fReadLibreFile(excelFilePath,excelSheetName)	
	Wscript.echo "~~~GINGER_RC_START~~~" 
	StdOut.Write("Outputvalue =")
	Wscript.StdOut.Write(ReadData)	
	Wscript.echo "~~~GINGER_RC_END~~~"
	
	
 Else
	ReadData2 = fReadDatafromDataFile(excelFilePath)
	Wscript.echo "~~~GINGER_RC_START~~~" 
	StdOut.Write("Outputvalue =")
	Wscript.StdOut.Write(ReadData2)	
	Wscript.echo "~~~GINGER_RC_END~~~"	
 End If


