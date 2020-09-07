'GINGER_Description Renamefile
'GINGER_$excelFilePath
'GINGER_$excelSheetName
'GINGER_$QC_URL
'GINGER_$QC_USERNAME
'GINGER_$QC_PASSWORD
'GINGER_$QC_DOMAIN
'GINGER_$QC_PROJECT
'GINGER_$QC_PATH
'GINGER_$QC_AutoScanParam_REQUIRED

'Option Explicit  'Line 10

if WScript.Arguments.Count = 0 then
    WScript.Echo "Missing parameters"        
end if

' Your code here
Dim excelFilePath
Dim excelSheetName
Dim strValue
Dim SNO
Dim FoundCell 
Dim Data_Var,FoundCell_COLUMN

excelFilePath = "C:\Ginger-Framework-Solution\Documents\1802\DATA_FILES_PER_CALENDAR\DEVELOPMENT\PRD1\DESHPANK\GOOGLE_SEARCH_CALENDER\GOOGLE_SEARCH_CALENDER.xlsx"
excelSheetName = "MAIN"
qcServer = "http://qc11isr1srv:8080/qcbin/start_a.jsp"
qcUser = "calendar"
qcPassword = "Export2QC"
qcDomain = "ATS_TESTING_ENV"
qcProject = "ABP_ DEV _TEST_92"
qcPath = "Root/New"
'qcTestSetName = ""
'qcTCName = WScript.Arguments(7)
'qcActivityName = WScript.Arguments(8)
'qcTCStatus = WScript.Arguments(9)
qcAutoScanParam = "Y"
'qcIntegrationRequired = WScript.Arguments(11)


If qcAutoScanParam = "Y" Then
                             'msgbox "Reached here"
                             Set tdc = CreateObject("TDApiOle80.TDConnection")      

                             strQCConnection = makeConnection(qcServer,qcUser,qcPassword,qcDomain,qcProject)

                             If strQCConnection = true Then
                                           'msgbox tdc.Connected
                                           Call fReadDatafromDataFile(excelFilePath)
                                           ReadData = "TC updated"
                             Else
                                           ReadData = "TC updation Failed."
                             End If

             

End If

Function makeConnection(qcServer,qcUser,qcPassword,qcDomain,qcProject)
              
              boolConnected = false
              'Set tdc = CreateObject("TDApiOle80.TDConnection")     
              
              tdc.InitConnectionEx qcServer
              
              If tdc.Connected Then
                             
                             'Login to QC
                             
                             tdc.Login qcUser,qcPassword
                             
                             If tdc.LoggedIn Then
                                           
                                           'Connect to QC Project
                                           tdc.Connect qcDomain, qcProject
                                           
                                           If tdc.ProjectConnected Then
                                                          boolConnected = true
                                                          'makeConnection = tdc
                                                          
                                           Else
                                                          boolConnected = false
                                           End If
                             Else
                                                          boolConnected = false   
                             End If
              Else 
                             boolConnected = false
              End If
              
              If boolConnected = true Then
                             makeConnection = true
              Else 
                             makeConnection = false
              End If
              
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
	
	Set FoundCell_COLUMN = objWSheet.Range("A1:AZ1").Find("TEST_NAME")
	Set FoundCell_Status_COLUMN = objWSheet.Range("A1:AZ1").Find("SKIP")
	
	If Not FoundCell_COLUMN Is Nothing Then
		Wscript.Echo "PARAM01 Doesn't found!"
	End If

        
	Set FoundCell = objWSheet.Range("B1:B"&strValue).Find("")

	If Not FoundCell Is Nothing Then

		'SNO = objWSheet.Cells(FoundCell.Row,1)
	  'Msgbox FoundCell.Row
		
       	  For j = 3  To 500

                If Trim(objWSheet.Cells(j, FoundCell_COLUMN.Column))<>"" then
		  If Trim(objWSheet.Cells(j, FoundCell_COLUMN.Column))<>"TEST_NAME" Then
			
			Conc_cnt=Trim(objWSheet.Cells(j, FoundCell_COLUMN.Column))
			objWSheet.Cells(j, FoundCell_Status_COLUMN.Column).Value= "Y"
		Else		
			Conc_cnt=Trim(objWSheet.Cells(j+1, FoundCell_COLUMN.Column))
			objWSheet.Cells(j+1, FoundCell_Status_COLUMN.Column).Value= "Y"
			
			
		 End If	
		 
		 Call GetALMTestSet(tdc, qcPath, Conc_cnt)
					   	
	           Data_Var = Data_Var & "," & Conc_cnt & ","
			

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
	Set FoundCell = Nothing
	set Copyrange = Nothing

fReadDatafromDataFile = SNO
 End Function
 
Public Function GetALMTestSet(strQCConnection,strFolderPath, strTestSetName)

	Set objTSTreeManager = strQCConnection.TestSetTreeManager
	Set objTSFolder = objTSTreeManager.NodeByPath(strFolderPath)	
	Set objTSList = objTSFolder.FindTestSets(strTestSetName)	
	
	boolTSFound = false
	If not objTSList Is Nothing Then
		'Search is success
		For each objTestSet in objTSList
			If objTestSet.Name = strTestSetName Then
				boolTSFound = true
				Exit For
			End If
		Next
	End If
	
	'If doesn't exist, create one
	If Not boolTSFound Then
		Set objTSFactory = objTSFolder.TestSetFactory
		Set objTestSet = objTSFactory.AddItem(Null)
		objTestSet.Name = strTestSetName
		objTestSet.Post
	End If

	Set GetALMTestSet = objTestSet

End Function
 
' output results to Ginger
Wscript.echo "~~~GINGER_RC_START~~~" 
WScript.Echo ReadData 
'Wscript.echo strVariable & "=" + ReadData 
Wscript.echo "~~~GINGER_RC_END~~~"