'o ########################################################################################
'Title:  Main file
'
' Author: Jayateerth
'
' Creation Date: 31-Oct-2016
'
' Description:  This script initilises the required functions and files required for automation set
'
'
' Called Scripts Used:
'
' Object Repository: Empty
'
' Test Prerequisites: N/A
'
' Data Prerequisites: N/A
'
' System Requirements: 
'
' Revision History: N/A
'
' Data Tables:  


'########################################################################################
'Global varibales declaration
'Test suite file Path

strTestSuite="E:\EN_Automation\EN_TestSuite.xls"

	

Public strStatus 		'Status Varibale for Continuation of execution
Public strAppUrl		'APplication URL

intRowCount   = 28		'Starting Row Number to Start Read Application Keyword from Test Suite
strStatus=True
Public strSummaryFilePath
Public strTCPath
Public strTestCaseName
Public strFeatureName
Public strKeyword
Public HTMLStepNo
Public strFilePath     'For reporting
Public TestCount   'Number of test cases executed
Public SummaryPassedStep
Public SummaryFailedStep
Public PassedStep
Public FailedStep
Public SummaryVerifyedStep
Public VerifiedStep
Public strDatabase
Public PassedTCCount
Public FailedTCCount
Public VerifiedTCCount


'Load Test suite
datatable.Import(strTestSuite)

strAppUrl=datatable.GetSheet("Global").GetParameter("parm_001").ValueByRow(8)
strBuild=datatable.GetSheet("Global").GetParameter("parm_001").ValueByRow(16)
strTCPath=datatable.GetSheet("Global").GetParameter("parm_001").ValueByRow(11)
strSummaryFilePath=datatable.GetSheet("Global").GetParameter("parm_001").ValueByRow(12)
'strSummaryFilePath=datatable.GetSheet("Global").GetParameter("parm_001").ValueByRow(13)
strDatabase=datatable.GetSheet("Global").GetParameter("parm_001").ValueByRow(17)

strRepo=datatable.GetSheet("Global").GetParameter("parm_001").ValueByRow(20) 'Repository
strLibFiles=datatable.GetSheet("Global").GetParameter("parm_001").ValueByRow(21) 'Lib files


'Associate the Repo
RepositoriesCollection.Add (strRepo)

'Split the Files path based on the semicolon
strLibArray=Split(strLibFiles,";")

For i = LBound(strLibArray) to UBound(strLibArray)
       
   ExecuteFile strLibArray(i)
   
   wait 2
   
   'Load the Function librarires
   LoadFunctionLibrary (strLibArray(i))
      
Next

'___________________________________________________________________________________
'Setup reporting files

Call SynchWait(2)


'Create Execution summary report file
CreateNewSummaryReportfile


'________________________________________________________________________________
'*******
'Exceute the Test cases based on TestSuite.xls
Call TestCaseExecute()

datatable.Export(strTestSuite)

'___________________________________________________________________
'Function to iterate through the Testsuite.xls
Public Function TestCaseExecute

TestCount=0
PassedTCCount=0
FailedTCCount=0
VerifiedTCCount=0

Do While strStatus = True

	
	
	strKeyword = datatable.GetSheet("Global").GetParameter("keyword").ValueByRow(intRowCount)


	a=StrComp(strKeyword,"EXEC", 1)   
	b=StrComp(strKeyword,"EXIT", 1)   
	If a = "0"  Then
	
	

		'strTCPath			 = datatable.GetSheet("Global").GetParameter("parm_001").ValueByRow(intRowCount)
		strTestCaseName		 = datatable.GetSheet("Global").GetParameter("comment").ValueByRow(intRowCount)
		strFeatureName		 = datatable.GetSheet("Global").GetParameter("parm_002").ValueByRow(intRowCount)
		strRunStauts		 = datatable.GetSheet("Global").GetParameter("parm_003").ValueByRow(intRowCount)
				
		strTestCase="TS_"&strTestCaseName&""
		
				
		If strFeatureName <> ""  Then
		'If strFeatureName <> ""  AND strRunStauts<> "PASS" Then   'to be implemented later
		
			'_________________________________
			'Repoting section
			repoPath="E:\EN_Automation\Reports\"&strTestCase&""
	
			screenShotPath="E:\EN_Automation\Reports\ScreenShots"
			
			
			'Generate Folders in local machine
			Set objShell = CreateObject("Wscript.Shell")
			objShell.Run "cmd /c mkdir "&repoPath&"\ScreenShots"
			
			
			'strFilePath =""&repoPath&"Report.html"
			strFilePath =""&repoPath&"\Report.html"
			
					
			'Create new Report HTML file
			CreateANewReportfile
			
					
			Call SynchWait(5)
			
			
		'________________________________
		
				'Execute the required Test script file
			'ExecuteFile "&strTCPath&"
			Call TestCaseExecution(strFeatureName,strTestCaseName,strTCPath)
			
				'Write Summary Test case Report
				WriteHTMLSummary
				
				'Set the Summary HTML Report  Test Case Status 
				call SummaryReportTCStatus
				
				'Write the Test case execution status to Test Suite
				If SummaryReportResult="PASS" Then
					DataTable.SetCurrentRow(intRowCount)
				
					DataTable.Value("parm_003", Global)="PASS"		
					PassedTCCount=PassedTCCount+1
									
				Else	
				
					DataTable.Value("parm_003", Global)="FAIL"
					FailedTCCount=FailedTCCount+1
				End If
				
'				Write Execution summary report
		SummaryPassedStep = SummaryPassedStep + PassedStep
		SummaryFailedStep = SummaryFailedStep + FailedStep
		SummaryVerifyedStep=SummaryVerifyedStep + VerifiedStep'Jay
		
		        
        ' Summary HTML Report - Write HTML Section
       WriteSummaryHTMLSection  SummaryHTMLStepNo,  strTestCaseName,  PassedStep,  FailedStep, VerifiedStep, SummaryReportResult, strFilePath'strFilePath'Jay
	
		TestCount=TestCount+1
		
		
		End If
		
		intRowCount = intRowCount + 1
		'TestCount  =TestCount  +1
		
		'Display the message for the user to set Testsuite.xls and terminate run
'		Msgbox "Please check Testsuite.xls and make sure data set is as expected.Test will terminate now"
'		ExitRun(0)
						
				
	ElseIf b="0"  then

					strStatus = False

	Else 

				intRowCount = intRowCount + 1

	End If


Loop


Call SynchWait(5)


''Write Execution summary report
WriteSummaryHTMLSummary


End Function

'_______________________________________________________