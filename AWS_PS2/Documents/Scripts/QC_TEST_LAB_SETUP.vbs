'GINGER_Description Upload Test Cases from Calendar to Test Plan and Test Lab
'GINGER_$CALENDAR_PATH
'GINGER_$CLENDAR_SHEET_NAME
'GINGER_$QC_URL
'GINGER_$QC_USERNAME
'GINGER_$QC_PASSWORD
'GINGER_$QC_DOMAIN
'GINGER_$QC_PROJECT
'GINGER_$QC_TEST_PLAN_PATH
'GINGER_$QC_TEST_LAB_PATH
'GINGER_$QC_INTEGRATION_REQUIRED
'GINGER_$QC_UPLOAD_ALL

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
Dim TestFound
Dim Testid

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
qcAutoScanParam = "N"
'qcIntegrationRequired = WScript.Arguments(11)

excelFilePath = WScript.Arguments(0)
excelSheetName = WScript.Arguments(1)
qcServer = WScript.Arguments(2)
qcUser = WScript.Arguments(3)
qcPassword = WScript.Arguments(4)
qcDomain = WScript.Arguments(5)
qcProject = WScript.Arguments(6)
qcTestPlanPath = WScript.Arguments(7)
qcTestLabPath = WScript.Arguments(8)
strQCIntergationReq = WScript.Arguments(9)
strQCUploadAll = WScript.Arguments(10)

'Added by Subodh Singh on 29-Aug-2018
'Create Test Plan and Test Lab path
'Condition - Test plan and Test plan folder should be same
'============================
If strQCIntergationReq = "Y" Then

			Set tdc = CreateObject("TDApiOle80.TDConnection")

            strQCConnection = makeConnection(qcServer,qcUser,qcPassword,qcDomain,qcProject)

			'Check folder exist in Test Lab
			Set objTSTreeManager = tdc.TestSetTreeManager
			Set objTSFolder = objTSTreeManager.NodeByPath(qcTestLabPath)

			'Check folder exist in Test Plan
			Set TestFact = tdc.TestFactory
			Set objTPTreeManager = tdc.TreeManager
			Set objTPFolder =objTPTreeManager.NodeByPath(qcTestPlanPath)

			'Read test case and test set name from calendar
			strTestSetName = fReadDatafromCalendar("TEST_SET_NAME")
			strTestCaseName = fReadDatafromCalendar("TEST_NAME")

			'Check TC exist in Test Plan 
			strTCCount = fnCheckTCExistinTetPlan(qcTestPlanPath, strTestCaseName)
			msgbox strTCCount
			
			'Check Test Case Already exist in given Test Set
			Call fnCheckTCExistinTestSet(qcTestLabPath, strTSName, strTSTCName)

            If strQCConnection = true Then
                'msgbox tdc.Connected
                'Call fReadDatafromDataFile(excelFilePath)
                ReadData = "TC updated"
            Else
                ReadData = "TC updation Failed."
            End If
End If
'=============================

If qcAutoScanParam = "Y" Then
                             'msgbox "Reached here"
                             Set tdc = CreateObject("TDApiOle80.TDConnection")

                             strQCConnection = makeConnection(qcServer,qcUser,qcPassword,qcDomain,qcProject)

                             If strQCConnection = true Then
                                           'msgbox tdc.Connected
                                           'Call fReadDatafromDataFile(excelFilePath)
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

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Function name: 	fReadDatafromCalendar()
' Description:  	Get PARAM_VALUE from calendar using PARAM_VALUE
' Parameters:   	PARAM_NAME(Method arguments)
'					excelFilePath, excelSheetName (Global Parameters)
' Return value:   	Success - True
' Failure - False
' Author:          	Subodh Singh
' Date:				31 AUG 2018
' Updated:
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Function fReadDatafromCalendar(PARAM_NAME)

	'Check File Exist
	If FileExist(excelFilePath) = False Then
		WScript.Echo "File Not Exist in "& excelFilePath
		fReadDatafromDataFile = "File Not Exist in "& excelFilePath
		Exit Function
	End If

	'Create Excel Objects
	Set objXls = CreateObject("Excel.Application")
	Set objWBook = objXls.Workbooks.Open(excelFilePath)
	Set objWSheet = objWBook.Worksheets(excelSheetName)

	'Check row exist with name "END" in column B
	Set FoundCell = objWSheet.Range("B1:B20000").Find("END")

	'Store last valid row number in variable - strValue
	If Not FoundCell Is Nothing Then
  		strValue =  FoundCell.Row
	End If

	'Store cell having empty SKIP column
    Set FoundCell_Empty = objWSheet.Range("B1:B" & strValue).Find("")

	On Error Resume Next

	'Store cell value having column name in found cells
	Set FoundCell = objWSheet.Range("A" & FoundCell_Empty.Row & ":BZ" & FoundCell_Empty.Row).Find(PARAM_NAME, , , 1)

	'Get PRARAM_VALUE for given column PARAM_NAME
	If Not FoundCell Is Nothing Then

		If Not FoundCell_Empty Is Nothing Then
			If PARAM_NAME="SKIP" Then
				objWSheet.Cells(FoundCell_Empty.Row,2).Value="X"
				objWSheet.Cells(FoundCell_Empty.Row+1,2).Value=PARAM_VALUE
			Else
				'msgbox "Row="&FoundCell.Row+1&",Col="&FoundCell.Column
				PARAM_VALUE = objWSheet.Cells(FoundCell.Row+1,FoundCell.Column).Value
			End If
		Else
			PARAM_VALUE = "none"
		End If
	End If

	'Save excel objects
	objWBook.Save
	objWBook.Close
	objXls.Quit

	'Clear excel objects
	Set objWSheet = Nothing
	Set objXls = Nothing
	Set FoundCell = Nothing
	set Copyrange = Nothing

	'Return PARAM_VALUE
	msgbox PARAM_VALUE
	fReadDatafromCalendar = PARAM_VALUE

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

Public Function fnCheckTCExistinTetPlan(strTPPath, strTCName)

	'Check folder exist in Test Plan
	Set TPTestFact = tdc.TestFactory
	Set objTPTreeManager = tdc.TreeManager
	Set objTPFolder =objTPTreeManager.NodeByPath(strTPPath)
	
	'Check if test case already exist in given test plan 
	strTestFactoryFilter = "select TS_TEST_ID from TEST where TS_NAME = '" & strTCName & "' and TS_SUBJECT = " & objTPFolder.NodeID
	Set objTPTestList = TPTestFact.NewList(strTestFactoryFilter)
	fnCheckTCExistinTetPlan = objTPTestList.Count
	
End Function

Public Function fnCheckTCExistinTestSet(strTSPath, strTSName, strTSTCName)
	'=================TO DO==========================
	'Check folder exist in Test Lab
	Set TSTestFact = tdc.TestSetFactory
	Set objTSTreeManager = tdc.TestSetTreeManager
	Set objTSFolder = objTSTreeManager.NodeByPath(strTSPath)
	
	If tsFolder Is Nothing Then  
        ReadData = "Path Not Found."
    Else
        'Msgbox "Path Found"
    End If
	
	' Search for the test set passed as an argument to the example code
    Set tsList = tsFolder.FindTestSets(strTSName)
    '----------------------------------Check if the Test Set Exists --------------------------------------------------------------------
    If tsList Is Nothing Then
        ReadData = "Test Set not found."
    End If
	
	Set theTestSet = tsList.Item(1)

        For Each testsetfound In tsList
              Set tsFolder = testsetfound.TestSetFolder
              Set tsTestFactory = testsetfound.tsTestFactory
              Set tsTestList = tsTestFactory.NewList("")

              For Each tsTest In tsTestList
				'MsgBox tsTest.Name
				  testrunname = "Test Case name"
				  If tsTest.Name = strTSTCName Then
					msgbox "Test Case Found"
				  End If
			  Next
		Next
	
	'fnCheckTCExistinTestSet = tsList.Count

End Function 

Public Function fnCreateTCinTestPlan(strTPPath,strTCName)
	
	'VerIfy If the test in the test plan otherwise it creates the test in the test plan 
	Set TestFact = tdc.TestFactory
	Set MyTMgr = tdc.TreeManager
	Set MySRoot = MyTMgr.NodeByPath(strTPPath) 
	
	'if main folder exists , main folder value will come from keep refer
	Set subjectfldr = MyTMgr.NodebyPath("Subject\" & folder)
	
	'subfolder value will come from calender
	'If subfolder = "" Then
      'Set trfolder = MyTMgr.NodebyPath("Subject\" & folder)
    'Else
      'Set trfolder = MyTMgr.NodebyPath("Subject\" & folder & "\" & subfolder)
    'End If
	
	Set objTestList = TestFact.NewList(strTCName)
	
	' if test case not found in test plan
	If objTestList.Count = 0 Then
		Set MyTest = trfolder.TestFact.AddItem(null)
		MyTest.Name = strTCName
		MyTest.Field("TS_SUBJECT") = MySRoot.NodeID
		
		MyTest.Post
		Testid = MyTest.id
		
	End If
	
	' if test case found in test plan
	If objTestList.Count > 0 Then
		TestFound = True 
		Set objTest = objTestList.Item(1)
		Testid = objTest.id
	End If
	
End Function 

Public Function fnCreateTCinTestSet(strTSPath,strTCName)

End Function 

' output results to Ginger
Wscript.echo "~~~GINGER_RC_START~~~"
WScript.Echo ReadData
'Wscript.echo strVariable & "=" + ReadData
Wscript.echo "~~~GINGER_RC_END~~~"
