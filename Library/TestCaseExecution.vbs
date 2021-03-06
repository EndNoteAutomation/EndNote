'**************************************************FrameWork Main Function*****************************************************************
'FunctionalityName: Functionality to Read Each Test Step of the Test Case from Excel
'BaseineVersion = V.0
'LastUpdated = 11-11-2016
'Developed By = Dilip Parashuram
Function TestCaseExecution(FeatureName,TestCaseExecuteName,sPath)
	
		
	Dim StepStatus,Status,Expected,Actual,TestStepStatus,PassCount,FailCount
	Set TestCaseExcel = Createobject("Excel.Application")
 	TestCaseExcel.Workbooks.Open (sPath)
	TestCaseExcel.Application.Visible = True
	Set FeatureNameSheet = TestCaseExcel.ActiveWorkbook.Worksheets(FeatureName)
	TestStepCount = 0
	TestStepRowCount = FeatureNameSheet.UsedRange.Rows.Count
	TestRowNumber = 0
	FailCount = 0
	PassCount = 0
	'Read Testcase Step from Test execution 
					For iTestStepRowCount = 2 To TestStepRowCount
						TestCaseName= FeatureNameSheet.Cells(iTestStepRowCount, 1).Value
						TestStepStatus = FeatureNameSheet.Cells(iTestStepRowCount, 6).Value
								If Trim(TestCaseName) = Trim(TestCaseExecuteName) Then
																					
										TestStepAction = FeatureNameSheet.Cells(iTestStepRowCount, 2).Value
										TestDataArray = FeatureNameSheet.Cells(iTestStepRowCount, 4).Value
										
										'Calling TestStepAction Function
										'Call GetTestStepObject(TestStepAction)
										
										StepStatus = Split(GetTestStepObject(TestStepAction)," - ")
										Expected = StepStatus(0)
										Actual = StepStatus(1)
										Status = StepStatus(2)
										
										
										If Ucase(Status) = "PASS" Then
											PassCount = PassCount+1
											
											FeatureNameSheet.Cells(iTestStepRowCount, 8).Interior.Color = vbGreen
											WriteHTMLSection  HTMLStepNo, Expected,Actual,Status," "
										Else
											FailCount = FailCount+1
											FeatureNameSheet.Cells(iTestStepRowCount, 8).Interior.Color = vbRED
											ScreenShotName = screenShotPath&"\"&TestCaseName&"_"&TestStepAction&".png"
											window("EndNote").CaptureBitmap(ScreenShotName),True
											WriteHTMLSection  HTMLStepNo, Expected,Actual,Status,ScreenShotName
											
										End If
										
										FeatureNameSheet.Cells(iTestStepRowCount, 8).Value = Status
										FeatureNameSheet.Cells(iTestStepRowCount, 7).Value = Actual
										FeatureNameSheet.Cells(iTestStepRowCount, 6).Value = Expected
										FeatureNameSheet.Cells(iTestStepRowCount, 9).Value = Now()
										TestStepCount = TestStepCount+1
								Else
										TestStepCount = TestStepCount+1
								End IF
					Next
	'End of Teststep reading 
	
	TestCaseExcel.ActiveWorkbook.Save
	TestCaseExcel.ActiveWorkbook.Close
	TestCaseExcel.Application.Quit
	Set TestCaseExcel = Nothing

	TestCaseExecution = PassCount&"-"&FailCount

End Function