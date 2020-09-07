'GINGER_Description QC Upload Script
'GINGER_$QC_URL
'GINGER_$QC_USERNAME
'GINGER_$QC_PASSWORD
'GINGER_$QC_DOMAIN
'GINGER_$QC_PROJECT
'GINGER_$QC_PATH
'GINGER_$QC_TEST_SET_NAME
'GINGER_$QC_TEST_CASE_NAME
'GINGER_$QC_GINGER_ACTIVITY_NAME
'GINGER_$QC_TEST_CASE_STATUS
'GINGER_$QC_UPLOAD_REQUIRED
'GINGER_$QC_INTEGRATION_REQUIRED




qcServer = WScript.Arguments(0)
qcUser = WScript.Arguments(1)
qcPassword = WScript.Arguments(2)
qcDomain = WScript.Arguments(3)
qcProject = WScript.Arguments(4)
qcPath = WScript.Arguments(5)
qcTestSetName = WScript.Arguments(6)
qcTCName = WScript.Arguments(7)
qcActivityName = WScript.Arguments(8)
qcTCStatus = WScript.Arguments(9)
qcUploadRequired = WScript.Arguments(10)
qcIntegrationRequired = WScript.Arguments(11)
'qcExecutionTime = WScript.Arguments(11)

If qcIntegrationRequired = "Y" Then
              If qcUploadRequired = 1 Then

                             'msgbox "Reached here"
                             Set tdc = CreateObject("TDApiOle80.TDConnection")      

                             strQCConnection = makeConnection(qcServer,qcUser,qcPassword,qcDomain,qcProject)

                             If strQCConnection = true Then
                                           'msgbox tdc.Connected
                                           Call UpdateQCStatus(tdc,qcPath,qcTCName,qcTCStatus)
                                           ReadData = "Status updated in QC."
                             Else
                                           ReadData = "QC Connection Failed."
                             End If

              End If
              
Else
              ReadData = "QC Integration Not Required."

End If

Function UpdateQCStatus(tdc,qcPath,qcTCName,qcTCStatus)

              Set TSetFact = tdc.TestSetFactory
              Set tsTreeMgr = tdc.testsettreemanager
              ' Get the test set folder passed as an argument to the example code
              nPath = Trim(qcPath)

              Set tsFolder = tsTreeMgr.NodeByPath(nPath)
              
              If tsFolder Is Nothing Then  
                             ReadData = "Path Not Found."
              Else
                             'Msgbox "Path Found"
              End If
              
              ' Search for the test set passed as an argument to the example code
              Set tsList = tsFolder.FindTestSets(qcTestSetName)
              '----------------------------------Check if the Test Set Exists --------------------------------------------------------------------
              If tsList Is Nothing Then
                             ReadData = "Test Set not found."
              End If

              '---------------------------------------------Check if the TestSetExists or is Duplicated ----------------------------------------------

              If tsList.Count > 1 Then
              ReadData = "FindTestSets found more than one test set: refine search"
              Exit Function
              ElseIf tsList.Count < 1 Then
              ReadData = "FindTestSets: test set not found"
              Exit Function
              End If

              '-------------------------------------------Access the Test Cases inside the Test SEt -------------------------------------------------

              Set theTestSet = tsList.Item(1)

              For Each testsetfound In tsList
              Set tsFolder = testsetfound.TestSetFolder
              Set tsTestFactory = testsetfound.tsTestFactory
              Set tsTestList = tsTestFactory.NewList("")

              For Each tsTest In tsTestList
              'MsgBox tsTest.Name
              testrunname = "Test Case name"
              If tsTest.Name = qcTCName Then

              '--------------------------------------------Accesss the Run Factory --------------------------------------------------------------------
                             Set RunFactory = tsTest.RunFactory
                             Set obj_theRun = RunFactory.AddItem(CStr(testrunname))
                             obj_theRun.Status = qcTCStatus '-- Status to be updated
                             obj_theRun.Post
                             
              '---------------------------------------------Update Run Step ----------------
                             Set oStep = obj_theRun.StepFactory
                             oStep.AddItem("Ginger Automation Execution")'Creating Step
                             Set oStepDetails = oStep.NewList("")
                             oStepDetails.Item(1).Field("ST_STATUS") = qcTCStatus'Updating Step Status
                             oStepDetails.Item(1).Field("ST_DESCRIPTION") = qcActivityName 'Updating Step Description
                             oStepDetails.Item(1).Field("ST_EXPECTED") = "Passed"'Updating Expected
                             oStepDetails.Item(1).Field("ST_ACTUAL") = qcTCStatus'Updating Actual
                             oStepDetails.Post
              End If
              Next 
                             Next 


End Function

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

' output results to Ginger
Wscript.echo "~~~GINGER_RC_START~~~" 
WScript.Echo ReadData 
'Wscript.echo strVariable & "=" + ReadData 
Wscript.echo "~~~GINGER_RC_END~~~"

