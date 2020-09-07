'GINGER_$ RUN_MODE
'GINGER_$ CALENDAR_NAME
'GINGER_$ VERSION
'GINGER_$ ENVIRONMENT
'GINGER_$ BF_NAME

Dim objFSO , objFile 
Dim Txt_File_Path 

sExecution_Mode = WScript.Arguments(0)
sExcelFile = WScript.Arguments(1)
'sORExcelfile = WScript.Arguments(2)
sRelease = WScript.Arguments(2)
' sEnvironment = WScript.Arguments(3)
sBusinessFlowName = WScript.Arguments(3)

Set objFSObject = CreateObject("Scripting.FileSystemObject")
vbsFullName = Wscript.ScriptFullName
vbsFile = objFSObject.GetFile(vbsFullName)
scriptsFullPath = objFSObject.GetParentFolderName(vbsFile) 
documentsFullPath=objFSObject.GetParentFolderName(scriptsFullPath)

'Path for Excels


'strExecutionpath = "D:\SVN\GingerSolutions\Ginger-ATT-HALO-GAMMA-TRINITY-SOR\Documents\TestData_XLS\" & sExecution_Mode & "\" & sRelease & "\Execution"
strExecutionpath = documentsFullPath & "\V" & sRelease & "\DATA_FILES_PER_CALENDAR\" & UCASE(sExecution_Mode) &"\" & Ucase(sEnvironment)
'strExecutionpath = documentsFullPath & "\TestData_XLS\" & sExecution_Mode & "\" & sRelease & "\Execution"
'strStoragepath = "D:\SVN\GingerSolutions\Ginger-ATT-HALO-GAMMA-TRINITY-SOR\Documents\TestData_XLS\" & sExecution_Mode & "\" & sRelease & "\Storage" 
'strStoragepath = documentsFullPath & "\TestData_XLS\" & sExecution_Mode & "\" & sRelease & "\Storage"   
strStoragepath = documentsFullPath & "\V" & sRelease & "\STORAGE\CALENDARS_MAIN_EXCEL\" & UCASE(sExecution_Mode)
strORStoragepath = documentsFullPath & "\V" & sRelease & "\STORAGE\OR"

'strExecutionpath = "D:\SVN\GingerSolutions\Ginger-ATT-HALO-GAMMA-TRINITY-SOR\Documents\TestData_XLS\" & sExecution_Mode & "\" & sRelease & "\Execution"
strExecutionpath = documentsFullPath & "\" & sRelease & "\DATA_FILES_PER_CALENDAR\" & UCASE(sExecution_Mode) &"\" & Ucase(sEnvironment)
'strExecutionpath = documentsFullPath & "\TestData_XLS\" & sExecution_Mode & "\" & sRelease & "\Execution"
'strStoragepath = "D:\SVN\GingerSolutions\Ginger-ATT-HALO-GAMMA-TRINITY-SOR\Documents\TestData_XLS\" & sExecution_Mode & "\" & sRelease & "\Storage" 
'strStoragepath = documentsFullPath & "\TestData_XLS\" & sExecution_Mode & "\" & sRelease & "\Storage"   
strStoragepath = documentsFullPath & "\" & sRelease & "\STORAGE\CALENDARS_MAIN_EXCEL\" & UCASE(sExecution_Mode)
strORStoragepath = documentsFullPath & "\" & sRelease & "\STORAGE\OR"



strExecutionpath = documentsFullPath & "\" & sRelease & "\DATA_FILES_PER_CALENDAR\" & UCASE(sExecution_Mode) &"\" & Ucase(sEnvironment) 
strExecutionpath_new = "\DATA_FILES_PER_CALENDAR\" & UCASE(sExecution_Mode) &"\" & Ucase(sEnvironment)
strStoragepath = documentsFullPath & "\" & sRelease & "\STORAGE\CALENDARS_MAIN_EXCEL\" & UCASE(sExecution_Mode)
strORStoragepath = documentsFullPath & "\" & sRelease & "\STORAGE\OR"





'C:\SVN\GingerSolutions\Ginger-ATT-HALO-GAMMA-TRINITY-SOR
strXlsFileName =  sExcelFile 
strORxlsFileName = sORExcelfile
strCommonFileName = "COMMON.xls"

strxlsfilenamelength= Instr(strXlsFileName,".")
strXlsFolderName=Left(strXlsFileName,strxlsfilenamelength-1)
Dim xlsFlag:xlsFlag = 1
Dim xlsORFlag:xlsORFlag = 1
Dim xlsCOMMONFlag:xlsCOMMONFlag = 1
Dim ErrorCode:ErrorCode="Success"


'Get NTNET username
objSysName = CreateObject("WScript.Network").UserName  
Set objFS=CreateObject("Scripting.FilesystemObject")    

'Check Documents Files Per Calendar Folder exist or not. If folder does not exist create folder 
'msgbox documentsFullPath & "\" & sRelease & "\DATA_FILES_PER_CALENDAR\"
If objFS.FolderExists(documentsFullPath & "\" & sRelease & "\DATA_FILES_PER_CALENDAR\") = False Then
	'msgbox strExecutionpath
    Set objnewFolderDataFilesPerCalendar = objFs.CreateFolder(documentsFullPath & "\" & sRelease & "\DATA_FILES_PER_CALENDAR\")
	
End If

'Check if Development folder exist or not. It folder doesnot exist then create folder
If objFS.FolderExists(documentsFullPath & "\" & sRelease & "\DATA_FILES_PER_CALENDAR\"& UCASE(sExecution_Mode)) = False Then
	'msgbox strExecutionpath
    
	Set objnewFolderDevExecMode = objFs.CreateFolder(documentsFullPath & "\" & sRelease & "\DATA_FILES_PER_CALENDAR\" & UCASE(sExecution_Mode))
	
End If 

'Check If File Is Already placed at correct location  // updated by shridhar 'sBusinessFlowName
If (objFS.FileExists(strExecutionpath & "\" & Ucase(objSysName) & "\" & Ucase(strXlsFolderName) & "\" & Ucase(sBusinessFlowName) & "\" & strXlsFileName)) Then 

 xlsFlag=0
End If

'Check If File Is Already placed at correct location  // updated by shridhar 'sBusinessFlowName
If (objFS.FileExists(strExecutionpath & "\" & Ucase(objSysName) & "\" & Ucase(strXlsFolderName) & "\" & Ucase(sBusinessFlowName) & "\" & strCommonFileName)) Then 
 xlsCOMMONFlag=0
End If

'Check If OR File Is Already placed at correct location 
'If (objFS.FileExists(strExecutionpath & "\" & Ucase(objSysName) & "\" & Ucase(strXlsFolderName) & "\" & strORxlsFileName)) Then 
 '		objFS.DeleteFile(strExecutionpath & "\" & Ucase(objSysName) & "\" & Ucase(strXlsFolderName) & "\" & strORxlsFileName)
'End If

'Verify if the Execution Folder Exist or not , if not create the folders .
arr_EXE_Folder = split(strExecutionpath_new ,"\")
arr_Count= UBound(arr_EXE_Folder )
strExecutionpath_new = documentsFullPath & "\" & sRelease & "\"
For i=0 to arr_Count 
	
	strExecutionpath_new = strExecutionpath_new & arr_EXE_Folder(i) &"\"			
	If objFS.FolderExists(strExecutionpath_new) = False Then 
		objFS.CreateFolder(strExecutionpath_new)
	End If
Next

'Verify Execution Folder exist in Development/Deployment folder or not. If it's not exist create New folder with system name
If objFS.FolderExists(strExecutionpath) = False Then
	'msgbox strExecutionpath
    Set objnewFolder = objFs.CreateFolder(strExecutionpath)   
End If

'Verify System Folder exist in Execution folder or not. If it's not exist create New folder with system name
If objFS.FolderExists(strExecutionpath &"\"& Ucase(objSysName)) = False Then
    Set objnewFolder = objFs.CreateFolder(strExecutionpath & "\" & Ucase(objSysName))   
End If
         
'Verify Excel file exist in Storage folder, if it not exist then stop execution
If objFS.FileExists(strStoragepath &"\"& strXlsFileName) = False Then        
    xlsFlag=0
    ScriptStatus = strXlsFileName & "Not Exist in Storage Folder" & strStoragepath     
    ErrorCode = "Excel File Not found in Storage !!" 
End If

'Verify COMMON Excel file exist in Storage folder, if it not exist then stop execution
If objFS.FileExists(strStoragepath &"\"& strCommonFileName) = False Then        
    xlsFlag=0
    ScriptStatus = strCommonFileName & "Not Exist in Storage Folder" & strStoragepath     
    ErrorCode = "Excel File Not found in Storage !!" 
End If

'Verify OR Excel file exist in Storage folder, if it not exist then stop execution
'If objFS.FileExists(strORStoragepath &"\"& strORxlsFileName) = False Then        
   ' xlsORFlag=0
  '  ScriptStatus = strORxlsFileName & "Not Exist in Storage Folder" & strStoragepath     
 '   ErrorCode = "OR Excel File Not found in Storage !!" 
'End If

'Create a Destination path inside the strExecutionpath >> Ntnet User >> Excel File Named Folder  updated by shridhar
strDestinationLocation_0 = strExecutionpath & "\" & Ucase(objSysName) & "\" & Ucase(strXlsFolderName)

'Create a Destination path inside the strExecutionpath >> Ntnet User >> Excel File Named Folder  updated by shridhar
strDestinationLocation = strExecutionpath & "\" & Ucase(objSysName) & "\" & Ucase(strXlsFolderName) & "\" & Ucase(sBusinessFlowName)


If xlsFlag=1 Then     

	'Verify destination folder exist in Execution folder or not, if it not exist then create a new folder 
    If objFS.FolderExists(strDestinationLocation_0) = False Then    	
        Set objnewFolder = objFs.CreateFolder(strDestinationLocation_0) 		
    End If 

    'Verify destination folder exist in Execution folder or not, if it not exist then create a new folder 
    If objFS.FolderExists(strDestinationLocation) = False Then		
        Set objnewFolder = objFs.CreateFolder(strDestinationLocation) 		
    End If  
	
    
    strSourcexlsLocation = strStoragepath & "\" & strXlsFileName  
    'strORSourcexlsLocation = strORStoragepath & "\" & strXlsFileName  
   
     
    If Right(strDestinationLocation,1) <> "\" Then
        strDestinationLocation = strDestinationLocation & "\"
    End If    
	
        
    'Copy the file from source to destination
    If objFS.FileExists(strDestinationLocation &"\"& strXlsFileName) = False Then    

		objFs.CopyFile strSourcexlsLocation, strDestinationLocation,False
    Else
	       ErrorCode = "File not copied"
    End If
	
	       
End If


If xlsFlag=0 Then     

	'Verify destination folder exist in Execution folder or not, if it not exist then create a new folder 
    If objFS.FolderExists(strDestinationLocation_0) = False Then    
	
        Set objnewFolder = objFs.CreateFolder(strDestinationLocation_0) 		
    End If 

    'Verify destination folder exist in Execution folder or not, if it not exist then create a new folder 
    If objFS.FolderExists(strDestinationLocation) = False Then
	
        Set objnewFolder = objFs.CreateFolder(strDestinationLocation) 		
    End If  

    strSourcexlsLocation = strStoragepath & "\" & strXlsFileName  
    'strORSourcexlsLocation = strORStoragepath & "\" & strXlsFileName  

     
    If Right(strDestinationLocation,1) <> "\" Then
        strDestinationLocation = strDestinationLocation & "\"
    End If    
	
		
    'Copy the file from source to destination
    If objFS.FileExists(strDestinationLocation &"\"& strXlsFileName) = True Then    

        objFs.CopyFile strSourcexlsLocation, strDestinationLocation,True
    Else
	       ErrorCode = "File not copied"
    End If
	
	    
   
End If


If xlsCOMMONFlag = 1 Then
	
	 strCommonxlsLocation = strStoragepath & "\" & strCommonFileName  
	 
	  If Right(strDestinationLocation,1) <> "\" Then
        strDestinationLocation = strDestinationLocation & "\"
    End If  
    
	   'Copy the Common excel file from source to destination
    If objFS.FileExists(strDestinationLocation &"\"& strCommonFileName) = False Then  
    	     
        objFs.CopyFile strCommonxlsLocation, strDestinationLocation,False
	Else
        ErrorCode = "File not copied"
    End If
		
End If

'If xlsORFlag=1 Then        

    'Verify destination folder exist in Execution folder or not, if it not exist then create a new folder 
'    If objFS.FolderExists(strDestinationLocation) = False Then            
'        Set objnewFolder = objFs.CreateFolder(strDestinationLocation) 
'    End If    
    
'    strORSourcexlsLocation = strORStoragepath & "\" & strORxlsFileName    
    
'    If Right(strDestinationLocation,1) <> "\" Then
'        strDestinationLocation = strDestinationLocation & "\"
'    End If    
    
    'Copy the OR file from source to destination
    'If objFS.FileExists(strDestinationLocation &"\"& strORxlsFileName) = False Then        
       ' objFs.CopyFile strORSourcexlsLocation, strDestinationLocation,False
    'Else
       ' ErrorCode = "File not copied"
    'End If
    
'End If


'*********************************************************************
'Function : To Create a Text file and save the XLSfilePath value to it 
'Developer : Mishu 
'Date : 28th Mar 2018
'*********************************************************************
Txt_File_Path = documentsFullPath&"\Log_Folder_Path.txt"
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Check if the File Areday Exists
If objFSO.FileExists(Txt_File_Path) Then 
	objFSO.DeleteFile Txt_File_Path
End If  

	Set objFile = objFSO.CreateTextFile(Txt_File_Path,True)
	If Right(strDestinationLocation,1)="\" Then
		 XLSfilePath_new= strDestinationLocation & strXlsFileName 
	Else
		XLSfilePath_new= strDestinationLocation & "\" & strXlsFileName 
	End If
	objFile.Write XLSfilePath_new
	objFile.Close


Set objFSO = Nothing
Set objFile = Nothing

'*************************************************************************



Wscript.Echo "~~~GINGER_RC_START~~~"
Wscript.Echo "Run Status =" & ErrorCode

If Right(strDestinationLocation,1)="\" Then
	Wscript.Echo "XLSfilePath =" & strDestinationLocation & strXlsFileName 
Else
	Wscript.Echo "XLSfilePath =" & strDestinationLocation & "\" & strXlsFileName 
End If

'Wscript.Echo "ORfilePath =" & strORStoragepath

Wscript.Echo "XLSfilePath =" & strDestinationLocation & "\" & strXlsFileName 
Wscript.Echo "ORfilePath =" & strORStoragepath

Wscript.Echo "XLSfilePath =" & strDestinationLocation & "\" & strXlsFileName 


Wscript.Echo "~~~GINGER_RC_END~~~"