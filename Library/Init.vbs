'########################################################################################
'Title:  INI file
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

intRowCount   = 28		'Starting Row Number to Start Read Application Keyword from Test Suite
strStatus=True
Public strTCPath
Public strTestCaseName
Public strFeatureName
Public strKeyword


 ExecuteFile "E:\EN_Automation\Library\Common_Functions.vbs"
' ExecuteFile "E:\EN_Automation\Library\Common_Functions.vbs\EN_Library.vbs"
' ExecuteFile  "E:\EN_Automation\Library\Common_Functions.vbs\TestCaseExecution.vbs"
'___________________________________________________
'UFT settings-Associate Repository, add Recovery scenario, Associate LIbraries etc using AOM
'strLibrary1="[ALM\Resources] Resources\TestAutomation\EN_Version_X8\Library\Common_Functions.vbs"
'strLibrary2="[ALM\Resources] Resources\TestAutomation\EN_Version_X8\Library\Extrnal_Files"
'strRepo="[ALM\Resources] Resources\TestAutomation\EN_Version_X8\Repository\EndNoteX8.tsr"
'
Dim strLibrary1,strLibrary2,strLibrary3
' strLibrary1 = "E:\EN_Automation\Library\Common_Functions.vbs"
' strLibrary2 = "E:\EN_Automation\Library\Common_Functions.vbs\EN_Library.vbs"
' strLibrary3 = "E:\EN_Automation\Library\Common_Functions.vbs\TestCaseExecution.vbs"
' strRepo="E:\EN_Automation\Repository\EndNoteX8.tsr"

' LoadFunctionLibrary (strLibrary1)
' LoadFunctionLibrary (strLibrary2)
' LoadFunctionLibrary (strLibrary3)

' RepositoriesCollection.Add (strRepo)

'____________________________________________________
'Load the files section

'Load Test suite
datatable.Import(strTestSuite)
strBuild=datatable.GetSheet("Global").GetParameter("parm_001").ValueByRow(16)
msgbox strBuild

'___________________________________________________
'Exceute the Test cases based on TestSuite.xls

Call TestCaseExecute()


Public Function TestCaseExecute

Do While strStatus = True
	
	strKeyword = datatable.GetSheet("Global").GetParameter("keyword").ValueByRow(intRowCount)
	msgbox strKeyword
	


	a=StrComp(strKeyword,"EXEC", 1)   
	b=StrComp(strKeyword,"EXIT", 1)   
	If a = "0"  Then

		strTCPath			 = datatable.GetSheet("Global").GetParameter("parm_001").ValueByRow(intRowCount)
		strTestCaseName		 = datatable.GetSheet("Global").GetParameter("comment").ValueByRow(intRowCount)
		strFeatureName		 = datatable.GetSheet("Global").GetParameter("parm_002").ValueByRow(intRowCount)
		
		msgbox strTCPath
								
		If strTCPath <> ""  Then
			
			'Execute the required Test script file
			Executefile strTCPath
	
	
		End If
		
		intRowCount = intRowCount + 1
		TestCount  =TestCount  +1
				
				
	ElseIf b="0"  then

					strStatus = False

	Else 

				intRowCount = intRowCount + 1

	End If



Loop

End Function

'_______________________________________________________


'_________________________________
'Repoting Individual report and Summary Report