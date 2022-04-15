Public ADODBInstructionSheet,ADODBDataTable
Public RS_RunInstruction,RS_DataTable
Dim fso,strScript,strFrameWorkPath
Dim strSQLquery_RunInstruction


'Create ADODB object
Set ADODBRunManager=CreateObject("ADODB.Connection")
'Create Record Set
Set RS_RunInstruction = CreateObject("ADODB.Recordset")
'Get Framework Path Dynamically
Set fso = CreateObject("Scripting.FileSystemObject")
strScript = Environment.Value("TestDir")
strFrameWorkPath = fso.GetParentFolderName(strScript)
Environment.Value("EnvFrameWorkPath") = strFrameWorkPath

'Load Library file and object repository at run time
'LoadLibFileAndSOR strFrameWorkPath&"\FunctionLibrary","qfl"
'LoadLibFileAndSOR strFrameWorkPath&"\SharedObjectRepository","tsr"

'Get All Value from Data Table Global Sheet and store into Environment Variable
Set ADODBDataTable_Global=CreateObject("ADODB.Connection")
'Create Record Set
Set RS_DataTableGlobal = CreateObject("ADODB.Recordset")

With ADODBDataTable_Global
	.Provider = "Microsoft.ACE.OLEDB.12.0"
    .ConnectionString = "Data Source="& Environment.Value("EnvFrameWorkPath") & "\DataTable\DataTable.xlsx" &";Extended Properties=""Excel 12.0 Xml;HDR=YES"""
    .Open
End With
strSQLquery_DataTable_Global = "Select * from [DT_Global$]"
RS_DataTableGlobal.Open strSQLquery_DataTable_Global,ADODBDataTable_Global
RS_DataTableGlobal.MoveFirst
While (NOT(RS_DataTableGlobal.EOF))
	Environment.Value(RS_DataTableGlobal.Fields.Item(0))=Trim(RS_DataTableGlobal.Fields.Item(1))
	RS_DataTableGlobal.MoveNext
Wend
RS_DataTableGlobal.Close
ADODBDataTable_Global.Close


'Create Result Path Folder
If Environment.Value("ResultPath")="" or Isnull(Environment.Value("ResultPath")) Then
	Environment.Value("EnvResultPath")=strFrameWorkPath&"\Results"
Else
	Environment.Value("EnvResultPath")=Environment.Value("ResultPath")
End If

If Not (fso.FolderExists(Environment.Value("EnvResultPath"))) Then
	fso.CreateFolder(Environment.Value("EnvResultPath"))
End If

'Open the connection for InstructionSheet and DataTab
With ADODBRunManager
     .Provider = "Microsoft.ACE.OLEDB.12.0"
     .ConnectionString = "Data Source="& Environment.Value("EnvFrameWorkPath") & "\RunManager.xlsx" &";Extended Properties=""Excel 12.0 Xml;HDR=YES"""
     .Open
End With


'Create and Open Recordset for  RunInstruction
strSQLquery_RunInstruction = "SELECT * FROM [RunInstruction$]where Execution_Flag = 'Yes'"
RS_RunInstruction.Open strSQLquery_RunInstruction,ADODBRunManager
RS_RunInstruction.MoveFirst

'Create Result Folder for test cases
sDate = Replace(Date, "/", "-")
sTime = Replace(Time, ":", "-")
strfolResult_TC = "" & Environment.Value("EnvResultPath") & "\" & "RS" & "_" & sDate & "_" & sTime
strfolHTML = strfolResult_TC &"\HTMLReport"
strfolScreenshot = strfolResult_TC &"\Screenshot"
Environment.Value("EnvfolHTML")=strfolHTML
Environment.Value("EnvfloScreenshot")=strfolScreenshot
fso.CreateFolder(strfolResult_TC)
fso.CreateFolder(strfolHTML)
fso.CreateFolder(strfolScreenshot)

'Create Header File of the HTML Report of Test Case
 CreateSummeryHTMLReport_Header
 Environment.Value("EnvTotalTC_PassCount")=0
Environment.Value("EnvTotalTC_FailCount")=0

'For All cases whose Execution_Flag mark as "YES"	
While (NOT(RS_RunInstruction.EOF))
	strSL_NO = RS_RunInstruction.Fields.item("SL_NO")
	strAuto_TC_ID = RS_RunInstruction.Fields.item("Auto_TC_ID")
	strTC_Description = RS_RunInstruction.Fields.item("TC_Description")
	strIteration = RS_RunInstruction.Fields.item("Iteration")
	strInstructionSheet = RS_RunInstruction.Fields.item("InstructionSheet")
	strDataTableSheet = RS_RunInstruction.Fields.item("DataTableSheet")
	strDataSequence_No = RS_RunInstruction.Fields.item("DataSequence_No")
	Environment.Value("EnvSLNo_TC")=strSL_NO
	Environment.Value("EnvAuto_TC_ID")=strAuto_TC_ID
	Environment.Value("EnvTC_Description")=strTC_Description
	Environment.Value("EnvTC_InstructionSheet")=strInstructionSheet
	Environment.Value("EnvTC_DataTableSheet")=strDataTableSheet
	Environment.Value("EnvTC_DataSequence_No")=strDataSequence_No
	
	'Calculate Iteration Number
	If Ucase(Trim(strIteration))="NO" Then
		intStartIteration=1
		intEndIteration=1
	ElseIf IsNumeric(Trim(strIteration)) Then
		intStartIteration=strIteration
		intEndIteration=strIteration
	ElseIf Ucase(Trim(strIteration))<>"NO" and Trim(strIteration)<>"" Then
		arrIterationNo=Split(Trim(strIteration),"-")
		intStartIteration=cint(arrIterationNo(0))
		intEndIteration=cint(arrIterationNo(1))
	End If
	
  	
  	'Create Header File of the HTML Report of Test Case
  	CreateHTMLReport_Header
  	
  	Environment.Value("EnvSLNo_ResultStep")=0
  	Environment.Value("EnvPassCount")=0
  	Environment.Value("EnvFailCount")=0
'  	
	'Execute Test Cases based on Iteration
	For intLoop = intStartIteration To intEndIteration Step 1
		intCurrentIterationNumber=intLoop
		Environment.Value("EnvTC_CurrentIteration")=intCurrentIterationNumber
		Call ExecuteDriver(strAuto_TC_ID,strInstructionSheet,strDataTableSheet,strDataSequence_No,intCurrentIterationNumber)
	Next
	
	'Create Footer File of the HMTL Report of Test Case
	CreateHTMLReport_Footer
	
	'Create Summery HTML Report Body
	CreateSummeryHTMLReport_Body
	
	RS_RunInstruction.MoveNext
Wend

'Create Summery HTML Report Footer
CreateSummeryHTMLReport_Footer

RS_RunInstruction.Close

ADODBRunManager.Close

Function ExecuteDriver(strAuto_TC_ID,strInstructionSheet,strDataTableSheet,strDataSequence_No,intCurrentIterationNumber)
	'Create ADODB Connection
	Set ADODBInstructionSheet=CreateObject("ADODB.Connection") 
	'Create Record Set
	Set RS_InstructionSheet = CreateObject("ADODB.Recordset")

	With ADODBInstructionSheet
     	.Provider = "Microsoft.ACE.OLEDB.12.0"
    	.ConnectionString = "Data Source="& Environment.Value("EnvFrameWorkPath") & "\InstructionSheet\InstructionSheet.xlsx" &";Extended Properties=""Excel 12.0 Xml;HDR=YES"""
     	.Open
	End With
	
	'Open Recordset for  Instruction
	strSQLquery_InstructionFlow = "Select * from ["&strInstructionSheet&"$] where Auto_TC_ID = '" & strAuto_TC_ID & "' and (`Comment` is null or `Comment` ='') order by Line_Number"
	RS_InstructionSheet.Open strSQLquery_InstructionFlow,ADODBInstructionSheet
	RS_InstructionSheet.MoveFirst
	
	'Execute all function
	While (Not(RS_InstructionSheet.EOF))
		strFunctionName=RS_InstructionSheet.Fields.item("FunctionName")
		strData1=RS_InstructionSheet.Fields.item("Data1")
		strData2=RS_InstructionSheet.Fields.item("Data2")
		strData3=RS_InstructionSheet.Fields.item("Data3")
		strData4=RS_InstructionSheet.Fields.item("Data4")
		strData5=RS_InstructionSheet.Fields.item("Data5")
		strData6=RS_InstructionSheet.Fields.item("Data6")
		strData7=RS_InstructionSheet.Fields.item("Data7")
		If not (Trim(strData1)="" ) Then
 			If Left(strData1, 1) = "#"Then
  				strData1 = Right(strData1,(len(strData1)-1))
 			Else 
  				sExtrenalizedDataTableName = strData1
  				strData1 =GetData_DataTable(sExtrenalizedDataTableName)
 			End If
		End If
		If not (Trim(strData2)="" ) Then
 			If Left(strData2, 1) = "#"Then
  				strData2 = Right(strData2,(len(strData2)-1))
 			Else 
  				sExtrenalizedDataTableName = strData2
  				strData2 =GetData_DataTable(sExtrenalizedDataTableName)
 			End If
		End If
		If not (Trim(strData3)="" ) Then
 			If Left(strData3, 1) = "#"Then
  				strData3 = Right(strData3,(len(strData3)-1))
 			Else 
  				sExtrenalizedDataTableName = strData3
  				strData3 =GetData_DataTable(sExtrenalizedDataTableName)
 			End If
		End If
		If not (Trim(strData4)="" ) Then
 			If Left(strData4, 1) = "#"Then
  				strData4 = Right(strData4,(len(strData4)-1))
 			Else 
  				sExtrenalizedDataTableName = strData4
  				strData4 =GetData_DataTable(sExtrenalizedDataTableName)
 			End If
		End If
		If not (Trim(strData5)="" ) Then
 			If Left(strData5, 1) = "#"Then
  				strData5 = Right(strData5,(len(strData5)-1))
 			Else 
  				sExtrenalizedDataTableName = strData5
  				strData5 =GetData_DataTable(sExtrenalizedDataTableName)
 			End If
		End If
		If not (Trim(strData6)="" ) Then
 			If Left(strData6, 1) = "#"Then
  				strData6 = Right(strData6,(len(strData6)-1))
 			Else 
  				sExtrenalizedDataTableName = strData6
  				strData6 =GetData_DataTable(sExtrenalizedDataTableName)
 			End If
		End If
		If not (Trim(strData7)="" ) Then
 			If Left(strData7, 1) = "#"Then
  				strData7 = Right(strData7,(len(strData7)-1))
 			Else 
  				sExtrenalizedDataTableName = strData7
  				strData7 =GetData_DataTable(sExtrenalizedDataTableName)
 			End If
		End If
		If blnStopExecution=True Then
			Exit Function
		End If
		Call FunctionCalling(strFunctionName,strData1,strData2,strData3,strData4,strData5,strData6,strData7)
		RS_InstructionSheet.MoveNext
	Wend
	RS_InstructionSheet.Close
	ADODBInstructionSheet.Close
End Function

Function FunctionCalling(strFunctionName,strData1,strData2,strData3,strData4,strData5,strData6,strData7)
	'Calling Function
	Select Case strFunctionName
		Case "LunchMMT"
			Call LunchMMT(strData1)
		Case "BookFlightOneWay"
			Call BookFlightOneWay
	End Select
End Function

Public Function LoadLibFileAndSOR(sPath,sFileType)
    Dim iFiles:iFiles = 0 
    Dim iCount:iCount = 0
    'Creating a object to work with Files
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    'Creating a object to work with folders inside a given Folder Path
    Set objFolders = objFSO.GetFolder(sPath)
    'Creating a object to work with Files 
    Set objFiles = objFolders.Files
    'Taking count of files Inside a particular folder path
    For Each sFiles in objFiles
        'Getting the Extension of the file
        sExtn = objFSO.GetExtensionName(sFiles.Name)
        If Ucase(sExtn) = Ucase(Trim(sFileType)) Then
            If Ucase(sExtn)="TSR" Then
            	RepositoriesCollection.Add sPath&"\"& sFiles.Name
            	'print sPath&"\"& sFiles.Name
            ElseIf Ucase(sExtn)="QFL" Then
            	LoadFunctionLibrary sPath&"\"& sFiles.Name
            	'print sPath&"\"& sFiles.Name
            End If
        End If
    Next
End Function
