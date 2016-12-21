' To genereate a HTML  File and Read, Write, Append data to the file

'Create a New File 


Dim Result
Dim PassedStep 
Dim FailedStep 
Dim VerifiedStep
Public RStatus

'PassedStep = 0				
'FailedStep = 0             
'RStatus = "Pass"           

'Public strFilePath
Public strBuild
Public strDatabase
'Public strTestCaseName
'Public HTMLStepNo

'*****Summary report*******
Public SummaryHTMLStepNo
Public SummaryReportResult
Public SummaryResultP
Public SummaryResultF
Public SummaryResultV


'HTMLStepNo = 1				


Sub CreateANewReportfile

		   Dim fso, MyFile
		   HTMLStepNo = 1
		   PassedStep = 0
		   FailedStep = 0
		    VerifiedStep=0 
		   RStatus = "Pass"
		   SummaryResultP = ""
		   SummaryResultV = ""
		   SummaryResultF = ""
		   SummaryReportResult = ""	

   
		   Set fso = CreateObject("Scripting.FileSystemObject")
		   'Set MyFile = fso.CreateTextFile(Environment.Value("strFilePath"), True)
		   Set MyFile = fso.CreateTextFile(strFilePath, True) 'For testing
		  		
		   '  Report  Title & Header
		   
		   MyFile.WriteLine("<HTML><BODY>")
		   MyFile.WriteLine("<BR>")
		   MyFile.WriteLine("<HEADER1><CENTER><U><B><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=3>AUTOMATED TEST EXECUTION REPORT</CENTER></HEADER1></B></U><BR>")
		   MyFile.WriteLine("<TABLE COLS=4 WIDTH='100%' BORDER=1>")
		   MyFile.WriteLine("<BR>")
		   MyFile.WriteLine("<TR>")
		   MyFile.WriteLine("<TD WIDTH='25%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>Date</B></FONT></TD>")
		   MyFile.WriteLine("<TD WIDTH='20%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>Build/Version</B></FONT></TD>")
		   'MyFile.WriteLine("<TD WIDTH='20%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>Build</B></FONT></TD>")          
		   MyFile.WriteLine("<TD WIDTH='25%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>Region/DataBase</B></FONT></TD>")
		   'MyFile.WriteLine("<TD WIDTH='25%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>Region/DataBase</B></FONT></TD>")
		   MyFile.WriteLine("<TD WIDTH='30%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>Test Case Name</B></FONT></TD>")
		   MyFile.WriteLine("</TR>")
		  
		   '  Report  Header  -  Application Specific Table Data  
		   
			MyFile.WriteLine("<TR>")
			MyFile.WriteLine("<TD BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>" & Now & "</B></FONT></TD>")
			MyFile.WriteLine("<TD BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>" & strBuild & "</B></FONT></TD>")
			MyFile.WriteLine("<TD BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>" & strDatabase & "</B></FONT></TD>")
			MyFile.WriteLine("<TD BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>" & strTestCaseName & "</B></FONT></TD>")
			MyFile.WriteLine("</TR>")
			MyFile.WriteLine("</TABLE>")
			MyFile.WriteLine("<BR>")
		
		 ' To Insert the Header Structure for the Report
		
			MyFile.WriteLine("<TABLE COLS=4 WIDTH='100%' BORDER=1>")
			MyFile.WriteLine("<TR>")
			MyFile.WriteLine("<TD WIDTH='5%' BGCOLOR='LIGHTSTEELBLUE'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>Step Name</B></FONT></TD>")
			MyFile.WriteLine("<TD WIDTH='45%' BGCOLOR='LIGHTSTEELBLUE'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>Expected Result</B></FONT></TD>")
			MyFile.WriteLine("<TD WIDTH='40%' BGCOLOR='LIGHTSTEELBLUE'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>Actual Result</B></FONT></TD>")
			MyFile.WriteLine("<TD WIDTH='10%' BGCOLOR='LIGHTSTEELBLUE'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>Status</b></FONT></TD>")
			MyFile.WriteLine("</TR>")
			
			' Close the file
			 MyFile.Close

End Sub
'________________________________________________________________________________________________________________________________________________________________


' Report   Summary for individual test cases

Public Function WriteHTMLSummary
   Const ForAppending = 8
   Const TristateUseDefault = -2
   Dim fso, f, ts
   Set fso = CreateObject("Scripting.FileSystemObject")
   'Set f = fso.GetFile(Environment.Value("strFilePath"))
   Set f = fso.GetFile(strFilePath) 'For reporting
   Set ts = f.OpenAsTextStream(ForAppending, TristateUseDefault)

 '   Report - Summary Details
   
    ts.Write "<BR>"
	ts.Write "<TABLE COLS=3 WIDTH='30%' BORDER=1>"
	ts.Write "<TR>"
	ts.Write "<BR>"
	ts.Write "<BR>"
	ts.Write "<BR>"
	ts.Write "<HEADER1><LEFT><U><B><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2>TEST EXECUTION SUMMARY</LEFT></HEADER1></B></U>"
	ts.Write "<TD WIDTH='20%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><CENTER><B>Total No of Steps</B></FONT></TD>"
	ts.Write "<TD WIDTH='20%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><CENTER><B>Steps Passed</B></FONT></TD>"
	ts.Write "<TD WIDTH='20%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><CENTER><B>Steps Failed</B></FONT></TD>"
	ts.Write "<TD WIDTH='20%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><CENTER><B>Steps Verified</B></FONT></TD>"'Jay
	ts.Write "</TR>"
	ts.Write "<TR>"
	ts.Write "<TD BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><CENTER><B>" & (PassedStep + FailedStep+VerifiedStep) & "</B></FONT></TD>"'Jay
    ts.Write "<TD BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><CENTER><B>"& PassedStep &"</B></FONT></TD>"
    ts.Write "<TD BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><CENTER><B>"& FailedStep & "</B></FONT></TD>"
    ts.Write "<TD BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><CENTER><B>"& VerifiedStep & "</B></FONT></TD>"'jay
    
    ts.Write "</TR><BR>"
	ts.Write "</TABLE>"
	
	ts.Write "<BR>"
	ts.Write "<TABLE COLS=2 WIDTH='60%' BORDER=1>"
	ts.Write "<TR>"
	ts.Write "<BR>"
	ts.Write "<BR>"
	ts.Write "<HEADER1><LEFT><U><B><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2>LEGEND: Notes</LEFT></HEADER1></B></U>"
	ts.Write "<TD WIDTH='8%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><LEFT><B>VERIFY</B></FONT></TD>"
	ts.Write "<TD WIDTH='70%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><LEFT><B>Please verify the steps by clicking the screenshot Manually & Mark the Test case as Pass or Fail appropriately</B></FONT></TD>"
	ts.Write "</TR>"
	ts.Write "<TR>"
	ts.Write "<TD WIDTH='8%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><LEFT><B>PASS</B></FONT></TD>"
	ts.Write "<TD WIDTH='70%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><LEFT><B>Execution of the steps were successful</B></FONT></TD>"
	ts.Write "</TR>"
	ts.Write "<TR>"
	ts.Write "<TD WIDTH='8%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><LEFT><B>FAIL</B></FONT></TD>"
	ts.Write "<TD WIDTH='70%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><LEFT><B>Execution of the steps were unsuccessful</B></FONT></TD>"
	ts.Write "</TR><BR>"
	ts.Write "</TABLE>"
	
	ts.Close
	
End Function

'________________________________________________________________________________________________________________________________________________________________________________________________________________
' To keep appending the messages to the Report

Public Function WriteHTMLSection(StepName1, Description1, ExpectedResult1, Result1, strFilePath1)


   Const ForAppending = 8
   Const TristateUseDefault = -2
   Dim fso, f, ts
   Set fso = CreateObject("Scripting.FileSystemObject")
   'Set f = fso.GetFile(Environment.Value("strFilePath"))
   Set f = fso.GetFile(strFilePath)'for reporting
   Set ts = f.OpenAsTextStream(ForAppending, TristateUseDefault)

 '   Report - Step Name, Description, Expected Result, Result
   
    ts.Write "<TR>"
	ts.Write "<TD WIDTH='10%'BGCOLOR='Gainsboro'><FONT FACE='VERDANA' SIZE=2>" & StepName1 & "</FONT></TD>"
	ts.Write "<TD WIDTH='40%'BGCOLOR='Gainsboro'><FONT FACE='VERDANA' SIZE=2>" & Description1 & "</FONT></TD>"
    ts.Write "<TD WIDTH='50%'BGCOLOR='Gainsboro'><FONT FACE='VERDANA' SIZE=2>" & ExpectedResult1 & "</FONT></TD>"
	HTMLStepNo = HTMLStepNo + 1


	If Result1 = "PASS" Then
		ts.Write "<TD WIDTH='10%'><B><FONT FACE='VERDANA'  COLOR='ROYALBLUE' SIZE=2>" & Result1 & "</B></FONT></TD>"
        PassedStep = PassedStep + 1
        SummaryResultP = Result1
        
        ts.Close    
    Else
      If Result1 = "ABORT" Then
        ts.Write "<TD WIDTH='10%'><B><FONT FACE='VERDANA'  COLOR='#D00000' SIZE=2>" & "FAIL" & "                  " & "<A href= " & strFilePath1 & "><B><FONT FACE='VERDANA' COLOR='#D00000' SIZE=1 <BR> <BR> Click for Screenshot</A>" & "</B></B></FONT></TD>"
        FailedStep = FailedStep + 1
        SummaryResultF = "FAIL"
	    RStatus = "Fail"
		ts.Close

		' Write the HTML Summary 
	    'WriteHTMLSummary									

        ' Send the HTML Failed Report
		'call SendNotesMail(strMailTo, strMailCC)			

		' Exit from the Execution completely
	   	ExitRun(0) 
    Else
      	If Result1 = "VERIFY" Then
           ts.Write "<TD WIDTH='10%'><B><FONT FACE='VERDANA'  COLOR='#FBB917' SIZE=2>" & Result1 & "                  " & "<A href= " & strFilePath1 & "> <B><FONT FACE='VERDANA' COLOR='#D00000' SIZE=1> <BR> Click for Screenshot</A>" & "</B></B></FONT></TD>"
        	'PassedStep = PassedStep + 1
        	VerifiedStep = VerifiedStep + 1 'jay
        	SummaryResultV = Result1
			ts.Close	
		Else
	  	If Result1 = "FAIL" Then
			ts.Write "<TD WIDTH='10%'><B><FONT FACE='VERDANA'  COLOR='#D00000' SIZE=2>" & Result1 & "                  " & "<A href= " & strFilePath1 & "> <B><FONT FACE='VERDANA' COLOR='#D00000' SIZE=1> <BR> Click for Screenshot</A>" & "</B></B></FONT></TD>"
			FailedStep = FailedStep + 1
			SummaryResultF = Result1
			ts.Close	
    	End if
    
    	End If	
	  End If 
	   
	   	   
	End If
	
	
	
End Function
'_________________________________________________________________________________________________________________________________________________________________________________________________________________
'***Summary report****
Sub CreateNewSummaryReportfile
		   Dim summaryfso, summaryMyFile
		   SummaryHTMLStepNo = 1
		   SummaryPassStep = 0
		   SummaryFailStep = 0
		    SummaryVerifyStep=0 'Jay
		   SummaryRStatus = "Pass"
		   Set Summaryfso = CreateObject("Scripting.FileSystemObject")
		   Set SummaryMyFile = Summaryfso.CreateTextFile(strSummaryFilePath, True) ''Updated here
		   
		   'Set SummaryMyFile = Summaryfso.CreateTextFile("D:\EndNoteTesting\Reports\Summary.html", True) ''Updated here
		   
		 		
		   '  Report  Title & Header
		   
		   SummaryMyFile.WriteLine("<HTML><BODY>")
		   SummaryMyFile.WriteLine("<BR>")
		   SummaryMyFile.WriteLine("<HEADER1><CENTER><U><B><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=3>AUTOMATED TEST EXECUTION SUMMARY REPORT</CENTER></HEADER1></B></U><BR>")
		   SummaryMyFile.WriteLine("<TABLE COLS=3 WIDTH='100%' BORDER=1>")
		   SummaryMyFile.WriteLine("<BR>")
		   SummaryMyFile.WriteLine("<TR>")
		   SummaryMyFile.WriteLine("<TD WIDTH='25%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>Date</B></FONT></TD>")
		   SummaryMyFile.WriteLine("<TD WIDTH='40%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>Build/Version</B></FONT></TD>")
		   
		   SummaryMyFile.WriteLine("<TD WIDTH='35%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>Region/DataBase</B></FONT></TD>")
		   
		   'summaryMyFile.WriteLine("<TD WIDTH='30%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>Test Case Name</B></FONT></TD>")
		   SummaryMyFile.WriteLine("</TR>")
		  
		   '  Report  Header  -  Application Specific Table Data  
		   		   
			SummaryMyFile.WriteLine("<TR>")
			SummaryMyFile.WriteLine("<TD BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>" & Now & "</B></FONT></TD>")
			SummaryMyFile.WriteLine("<TD BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>" & strBuild & "</B></FONT></TD>")
			SummaryMyFile.WriteLine("<TD BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>" & strDatabase & "</B></FONT></TD>")
			'summaryMyFile.WriteLine("<TD BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>" & strTestCaseName & "</B></FONT></TD>")
			SummaryMyFile.WriteLine("</TR>")
			SummaryMyFile.WriteLine("</TABLE>")
			SummaryMyFile.WriteLine("<BR>")
		
		 ' To Insert the Header Structure for the Report
		
			SummaryMyFile.WriteLine("<TABLE COLS=6 WIDTH='100%' BORDER=1>")
			SummaryMyFile.WriteLine("<TR>")
			SummaryMyFile.WriteLine("<TD WIDTH='5%' BGCOLOR='LIGHTSTEELBLUE'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>Step Name</B></FONT></TD>")
			SummaryMyFile.WriteLine("<TD WIDTH='40%' BGCOLOR='LIGHTSTEELBLUE'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>Test Case Name</B></FONT></TD>")
			SummaryMyFile.WriteLine("<TD WIDTH='10%' BGCOLOR='LIGHTSTEELBLUE'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>No. of Steps</B></FONT></TD>")
			SummaryMyFile.WriteLine("<TD WIDTH='10%' BGCOLOR='LIGHTSTEELBLUE'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>Passed Steps</b></FONT></TD>")
			SummaryMyFile.WriteLine("<TD WIDTH='10%' BGCOLOR='LIGHTSTEELBLUE'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>Failed Steps</b></FONT></TD>")
			SummaryMyFile.WriteLine("<TD WIDTH='10%' BGCOLOR='LIGHTSTEELBLUE'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>Steps to Be Verified </b></FONT></TD>")'Jay
			SummaryMyFile.WriteLine("<TD WIDTH='15%' BGCOLOR='LIGHTSTEELBLUE'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><B>Status</b></FONT></TD>")
			SummaryMyFile.WriteLine("</TR>")
			
			' Close the file
			 SummaryMyFile.Close

End Sub
'_______________________________________________________________________________________________________________________________________________________________________________________
' Summary HTML Report Test Case status - Test Case status at the end of the Test case execution
Public Function SummaryReportTCStatus

	If SummaryResultP <> "" or SummaryResultV <> "" or SummaryResultF <> "" Then
	
		SummaryReportResult = "PASS"
		
		If SummaryResultV <> "" or SummaryResultF <> "" Then
		
			SummaryReportResult = "VERIFY"
			
			If SummaryResultF <> ""  Then
			
				SummaryReportResult = "FAIL"
				
			End If
			
		End If
		
	End If		

End function
'_______________________________________________________________________________________________________________________________________________________________________________________
' Summary - Report   Summary 

Public Function WriteSummaryHTMLSummary
   Const ForAppending = 8
   Const TristateUseDefault = -2
   Dim Summaryfso, Summaryf, Summaryts
   Set Summaryfso = CreateObject("Scripting.FileSystemObject")
'   Set Summaryf = Summaryfso.GetFile("D:\EndNoteTesting\Reports\Summary.html")
		
	Set Summaryf = Summaryfso.GetFile(strSummaryFilePath)
   Set Summaryts = Summaryf.OpenAsTextStream(ForAppending, TristateUseDefault)

 '   Report - Summary Details
         
    Summaryts.Write "<BR>"
	Summaryts.Write "<TABLE COLS=3 WIDTH='35%' BORDER=1>"
	Summaryts.Write "<TR>"
	Summaryts.Write "<BR>"
	Summaryts.Write "<BR>"
	Summaryts.Write "<BR>"
	Summaryts.Write "<HEADER1><LEFT><U><B><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2>TEST EXECUTION SUMMARY</LEFT></HEADER1></B></U>"
	Summaryts.Write "<TD WIDTH='20%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><CENTER><B>Total No. of Test Scripts</B></FONT></TD>"
	Summaryts.Write "<TD WIDTH='25%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><CENTER><B>Total No. of Passed Scripts</B></FONT></TD>"
	Summaryts.Write "<TD WIDTH='20%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><CENTER><B>Total No. of Failed Scripts</B></FONT></TD>"
	Summaryts.Write "<TD WIDTH='20%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><CENTER><B>Total No. of Scripts to be Verified </B></FONT></TD>"'Jay
	Summaryts.Write "</TR>"
	Summaryts.Write "<TR>"
	'Summaryts.Write "<TD BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><CENTER><B>" & (SummaryPassedStep +  SummaryFailedStep+SummaryVerifyedStep) & "</B></FONT></TD>"'Jay
    Summaryts.Write "<TD BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><CENTER><B>" & (TestCount) & "</B></FONT></TD>"'Jay
    
	Summaryts.Write "<TD BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><CENTER><B>"& PassedTCCount &"</B></FONT></TD>"
    Summaryts.Write "<TD BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><CENTER><B>"& FailedTCCount & "</B></FONT></TD>"
    Summaryts.Write "<TD BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><CENTER><B>"& VerifiedTCCount & "</B></FONT></TD>"'jay    
    
'    Summaryts.Write "<TD BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><CENTER><B>"& SummaryPassedStep &"</B></FONT></TD>"
'    Summaryts.Write "<TD BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><CENTER><B>"& SummaryFailedStep & "</B></FONT></TD>"
'    Summaryts.Write "<TD BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><CENTER><B>"& SummaryVerifyedStep & "</B></FONT></TD>"'jay
    Summaryts.Write "</TR><BR>"
	Summaryts.Write "</TABLE>"
	
	Summaryts.Write "<BR>"
	
	Summaryts.Write "<TABLE COLS=2 WIDTH='60%' BORDER=1>"
	Summaryts.Write "<TR>"
	Summaryts.Write "<BR>"
	Summaryts.Write "<BR>"
	Summaryts.Write "<HEADER1><LEFT><U><B><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2>LEGEND: Notes</LEFT></HEADER1></B></U>"
	Summaryts.Write "<TD WIDTH='8%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><LEFT><B>VERIFY</B></FONT></TD>"
	Summaryts.Write "<TD WIDTH='70%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><LEFT><B>Please verify the steps by clicking the screenshot Manually & Mark the Test case as Pass or Fail appropriately</B></FONT></TD>"
	Summaryts.Write "</TR>"
	Summaryts.Write "<TR>"
	Summaryts.Write "<TD WIDTH='8%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><LEFT><B>PASS</B></FONT></TD>"
	Summaryts.Write "<TD WIDTH='70%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><LEFT><B>Execution of the steps were successful</B></FONT></TD>"
	Summaryts.Write "</TR>"
	Summaryts.Write "<TR>"
	Summaryts.Write "<TD WIDTH='8%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><LEFT><B>FAIL</B></FONT></TD>"
	Summaryts.Write "<TD WIDTH='70%' BGCOLOR='Bisque'><FONT FACE='VERDANA' COLOR='SaddleBrown' SIZE=2><LEFT><B>Execution of the steps were unsuccessful</B></FONT></TD>"
	Summaryts.Write "</TR><BR>"
	Summaryts.Write "</TABLE>"
			
'	Set Summaryf=Nothing
'	Set Summaryts=Nothing
'	Set Summaryfso=Nothing
	Summaryts.Close
	
End Function
'______________________________________________________________________________________________________________________
' To keep appending the messages to the Summary Report

Public Function WriteSummaryHTMLSection(SummaryStepName1, strTestCaseName1, SummaryPassStep1, SummaryFailStep1,SummaryVerifyStep1, SummaryResult1, strSummaryFilePath1)'jay
   Const ForAppending = 8
   Const TristateUseDefault = -2
   Dim Summaryfso, Summaryf, Summaryts
   Set Summaryfso = CreateObject("Scripting.FileSystemObject")
'   Set Summaryf = Summaryfso.GetFile("D:\EndNoteTesting\Reports\Summary.html")
	Set Summaryf = Summaryfso.GetFile(strSummaryFilePath)
   Set Summaryts = Summaryf.OpenAsTextStream(ForAppending, TristateUseDefault)

 '   Report - Step Name, Test Case Name, Total No. of Steps, Total No. of Passed steps, Total No. of Failed steps, Result
     
    Summaryts.Write "<TR>"
	Summaryts.Write "<TD WIDTH='5%'BGCOLOR='Gainsboro'><FONT FACE='VERDANA' SIZE=2>" & SummaryStepName1 & "</FONT></TD>"
	Summaryts.Write "<TD WIDTH='40%'BGCOLOR='Gainsboro'><FONT FACE='VERDANA' SIZE=2>" & strTestCaseName1 & "</FONT></TD>"
    Summaryts.Write "<TD WIDTH='10%'BGCOLOR='Gainsboro'><FONT FACE='VERDANA' SIZE=2>" & SummaryPassStep1 + SummaryFailStep1+SummaryVerifyStep1 & "</FONT></TD>"'Jay
    Summaryts.Write "<TD WIDTH='10%'BGCOLOR='Gainsboro'><FONT FACE='VERDANA' SIZE=2>" & SummaryPassStep1 & "</FONT></TD>"
    Summaryts.Write "<TD WIDTH='10%'BGCOLOR='Gainsboro'><FONT FACE='VERDANA' SIZE=2>" & SummaryFailStep1 & "</FONT></TD>"
    Summaryts.Write "<TD WIDTH='10%'BGCOLOR='Gainsboro'><FONT FACE='VERDANA' SIZE=2>" & SummaryVerifyStep1 & "</FONT></TD>"'Jay
    SummaryHTMLStepNo = SummaryHTMLStepNo + 1

	If SummaryResult1 = "PASS" Then
	
		Summaryts.Write "<TD WIDTH='15%'><B><FONT FACE='VERDANA'  COLOR='ROYALBLUE' SIZE=2>" & SummaryResult1 & "                  " & "<A href= " & strSummaryFilePath1 & "> <B><FONT FACE='VERDANA' COLOR='ROYALBLUE' SIZE=1> <BR> Click for Detailed Report</A>" & "</B></B></FONT></TD>"
        SummaryPassStep = SummaryPassStep + 1
		Summaryts.Close    
    Else
      If SummaryResult1 = "ABORT" Then
        Summaryts.Write "<TD WIDTH='15%'><B><FONT FACE='VERDANA'  COLOR='#D00000' SIZE=2>" & "FAIL" & "                  " & "<A href= " & strSummaryFilePath1 & "><B><FONT FACE='VERDANA' COLOR='#D00000' SIZE=1 <BR> <BR> Click for Detailed Report</A>" & "</B></B></FONT></TD>"
        SummaryFailedStep = SummaryFailedStep + 1
	    SummaryRStatus = "Fail"
		Summaryts.Close

		' 	Write the Summary HTML Summary 
	   ' WriteSummaryHTMLSummary									

        ' 	Send the Summary HTML Failed Report
		'call SendSummaryNotesMail(strSummaryMailTo, strSummaryMailCC)			
		'call SendMail(strSummaryMailTo, strSummaryMailCC)	'Outlook Mail Pending
		
		'call OutlookmailBypassSec(strSummaryMailTo, strSummaryMailCC)
		
		'call SendMailSMTP (strSummaryMailTo, strSummaryMailCC)
		
		' 	Exit from the Execution completely
	   	ExitRun(0) 
    Else
      	If SummaryResult1 = "VERIFY" Then
           Summaryts.Write "<TD WIDTH='15%'><B><FONT FACE='VERDANA'  COLOR='#FBB917' SIZE=2>" & SummaryResult1 & "                  " & "<A href= " & strSummaryFilePath1 & "> <B><FONT FACE='VERDANA' COLOR='ROYALBLUE' SIZE=1> <BR> Click for Detailed Report</A>" & "</B></B></FONT></TD>"'Jay
           'SummaryPassStep = SummaryPassStep + 1
		   SummaryVerifyedStep=SummaryVerifyedStep+1
		   Summaryts.Close	
	Else
	  	If SummaryResult1 = "FAIL" Then
			Summaryts.Write "<TD WIDTH='15%'><B><FONT FACE='VERDANA'  COLOR='#D00000' SIZE=2>" & SummaryResult1 & "                  " & "<A href= " & strSummaryFilePath1 & "> <B><FONT FACE='VERDANA' COLOR='#D00000' SIZE=1> <BR> Click for Detailed Report</A>" & "</B></B></FONT></TD>"
			SummaryFailStep = SummaryFailStep + 1
			Summaryts.Close	
    	End if
    
    	End If	
	   End If 
	End If
	
	' HTML Report - Summary 
		'WriteSummaryHTMLSummary
End Function
'_________________________________________________________________________________________________________________________________________________________________________________________________________________
' Summary HTML Report - ALL PASS scenario - Has to be at the END of the Execution of the Integrated Script
' Sends Mail at the end of Execution
Private Function SummaryHTMLPass

 	If SummaryRStatus = "Pass" Then
 	
 		' HTML Report - Summary 
		WriteSummaryHTMLSummary

		' Send Outlook Mail   - Sends mail From Outlook with the HTML attachment
		'call OutlookmailBypassSec(strSummaryMailTo, strSummaryMailCC)
		
	
		
		'call SendMailSMTP (strSummaryMailTo, strSummaryMailCC)
		
	End if
	
End Function
'______________________________________________________________________________________________________________________________________________________________________________________
Public Sub SendEmail(strsubject,stremailcontent,arrAttachements,strrecipient)
Const olMailItem = 0

Set OAPp = CreateObject("Outlook.Application")
Set myItem = OApp.CreateItem(olMailItem)
With myItem
	.To = strrecipient
	.Subject = strsubject & " : : " & now
	.BodyFormat = 2
	.HtmlBody = stremailcontent
	
	'For each element in arrAttachements
	.Attachments.Add arrAttachements
	'Next
	
	wait 2 'adding attachements may take sometime
	
	.Display 'Display the email message but dont send otherwise security pop up will come up 
End with

Set oWin = Window("text:=Automated Email Notification : QTP Statement Script Execution Results : : SendEmail - .*")
oWin.Activate
SendKeys "%S"

Set myItem = Nothing
Set oWin = Nothing

End Sub
'__________________________________________________________________________________________________
Function Sendkeys(Keystroke)
Dim oShell,hwnd,title
If Keystroke ="" Then
	Exit Function
End If

Set oShell = CreateObject("WScript.Shell") 
Extern.Declare micHWnd, "GetForegroundWindow", "user32.dll", "GetForegroundWindow" 
hwnd = Extern.GetForegroundWindow()
title = Window("hwnd:=" & hwnd ).GetROProperty( "title" )
oShell.AppActivate title
oShell.SendKeys Keystroke
Wait 0, 500

Set oShell = Nothing
End Function
'________________________________________________________________________________________________