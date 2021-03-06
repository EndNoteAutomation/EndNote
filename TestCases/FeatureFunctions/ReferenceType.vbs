'**************************************
'functionality:Reference Type/Data
'purpose:default fields of Reference type
'Created by:Dilip Parashuram 
'created Date:29/08/2016
'modified by;
'Modified date:
'**************************************
Option Explicit 
Dim path,TesetStepExecuteName,oShell,TestDataArray,strAuthorName,strAuthWinId,authName,ReferenceTypeName,RefType,Iterator,ReftypeCount,ReferenceFields
Dim EMessage,ReferenceFields1,ReferenceFields2,ReferenceFields3,J
Set oShell =  CreateObject("WScript.shell")
'Load Library Functions & repository 
'Execution the Test Objects

Function GetTestStepObject(TestStepAction)

Dim Status, Actual, Expected
	Select Case TestStepAction
	
	
		Case "LaunchApplication"
			If Window("EndNote").Exist(3) = False Then
				call launchApplication()
			End If
			If Window("EndNote").Exist(5) = true Then
				Status = "PASS"
				Expected = "Endnote Must be Launched"
				Actual = "Endnote is Launched"
				else
				Status = "FAIL"
				Expected = "Endnote Must be Launched"
				Actual = "Endnote is Not Launched"
				
			End If
			
		Case "CloseLibraryFile"
			Call CloseLibraryFile()
			
			If NOT Window("EndNote").Window("New Reference window").Exist(2) = TRUE Then
				Status = "PASS"
				Expected = "All the Existing Windows must Close successfully"
				Actual = "All the Existing Windows are Closed successfully"
			else
				Status = "FAIL"
				Expected = "All the Existing Windows must Close successfully"
				Actual = "All the Existing Windows are NOT Closed successfully"
	
			End If
			
		Case "CloseReferenceFile"
			Call CloseReferenceFile()
			
			'If NOT Window("EndNote").Window("New Reference window").Exist(2) = TRUE Then
				Status = "PASS"
				Expected = "All the Existing Windows must Close successfully"
				Actual = "All the Existing Windows are Closed successfully"
'			else
'				Status = "FAIL"
'				Expected = "All the Existing Windows must Close successfully"
'				Actual = "All the Existing Windows are NOT Closed successfully"
'	
'			End If
				
		Case "OpenSampleLibrary"
			call OpenSampleLilbrary()
			If Window("EndNote").Exist(2) = true Then
				Status = "Pass"
				Expected = "Sample Library Must be Opened Successfully"
				Actual = "Sample Library is Opened Successfully"
				else
				Status = "FAIL"
				Expected = "Sample Library Must be Opened Successfully"
				Actual = "Sample Library is Not Opened Successfully"
				
			End If
		Case "OpenNewReferenceWindow"
			Window("EndNote").WinMenu("Menu").Select "References;New Reference	Ctrl+N"

			If Window("EndNote").Window("New Reference window").Exist(2) = TRUE Then
				Status = "PASS"
				Expected = "New Reference Window is Opened successfully"
				Actual = "New Referene Window Is Opeded Successfully"
			else
				Status = "FAIL"
				Expected = "New Reference Window is Opened successfully"
				Expected = "New Reference Window is NOT Opened successfully"
	
			End If
			
		Case "SelectReferenceType"
			If Window("EndNote").Dialog("EndNote Preferences").Exist(2) = TRUE Then
				Window("EndNote").Dialog("EndNote Preferences").WinComboBox("DefaultReferenceType").Select "Aggregated Database"
				oShell.SendKeys "ENTER"
			ElseIf Window("EndNote").Window("New Reference window").Exist(2) = TRUE Then
				Window("EndNote").InsightObject("ReferenceType").Click
				ReferenceTypeName = split(TestDataArray,"|")
				RefType = ReferenceTypeName(0)
				RefType = Left(RefType,1)
				ReftypeCount = ReferenceTypeName(1)
				For Iterator = 1 To ReftypeCount
					RefType = """"&RefType&""""
					oShell.SendKeys RefType
				Next
				oShell.SendKeys "{ENTER}"
				
				Status = "PASS"
				Expected = "ReferenceType "&ReferenceTypeName(0)&" Should be Selected Successfully"
				Actual = "ReferenceType Window is Selected Successfully"
			End If	
			

		Case "VerifyTheReferenceTypeFields"
			If Window("EndNote").Window("New Reference").Exist(2) = TRUE Then
			Window("EndNote").Window("New Reference").WinObject("Author").Click
			ReferenceFields = Window("EndNote").Window("New Reference").WinObject("NewReferenceFieldArea").GetVisibleText
			
			ReferenceTypeName = split(TestDataArray,"|")
				'RefType = ReferenceTypeName(0)
				ReftypeCount = ubound(ReferenceTypeName)
				For Iterator = 0 To ReftypeCount
					'RefType = """"&RefType&""""
					'If ReftypeCount = "3" Then
					For J = 1 To ReferenceTypeName(Iterator)
							oShell.SendKeys "{TAB}"	
					Next
						If Iterator = 0 Then
							ReferenceFields1 = Window("EndNote").Window("New Reference").WinObject("NewReferenceFieldArea").GetVisibleText
						ElseIf Iterator = 1 Then
							ReferenceFields2 = Window("EndNote").Window("New Reference").WinObject("NewReferenceFieldArea").GetVisibleText
						ElseIf Iterator = 2 Then	
							ReferenceFields3 = Window("EndNote").Window("New Reference").WinObject("NewReferenceFieldArea").GetVisibleText
						End If
				Next
			ReferenceFields = ReferenceFields&ReferenceFields1&ReferenceFields2
				If NOT  Trim(ReferenceFields) = "" Then
					Expected = "The Default Reference Type Fields should be Displayed"
					Actual = ReferenceFields&" are the reference Fields displayed"
					Status = "PASS"	
				Else
					Status = "FAIL"
					Expected = "The Default Reference Type Fields should be Displayed"
					Actual = "No Reference Fields are displayed"
				End IF
			End IF
		
		Case "CloseWarningWindow"
			IF Window("EndNote").Dialog("EndNote").Exist(2) = True THEN
					EMessage = Window("EndNote").Dialog("EndNote").Static("This Record is Empty").GetROProperty("text")
					Status = "PASS"
					Expected = "Warning message should be Displayed"
					Actual = EMessage&" message Displayed Successfully"
			else
				Status = "FAIL"
				Expected = "Warning message should be Displayed"
				Actual = EMessage&" is NOT Displayed Successfully"
			
			End IF
			Window("EndNote").Dialog("EndNote").WinButton("OK").Click
		
		Case "EN_Select_Reference"
		
			Window("EndNote").Dialog("EndNote Preferences").WinTreeView("PreferenceOption").Select "Reference Types"
			If Window("EndNote").Dialog("EndNote Preferences").WinComboBox("DefaultReferenceType").Exist(2) =  True Then
			Status = "PASS"
			Expected = "ReferenceType Window should be displayed Successfully"
			Actual = "Reference Type Window is Displayed Successfully"
			else
			Status = "FAIL"
			Expected = "ReferenceType Window should be displayed Successfully"
			Actual = "Reference Type Window is NOT Displayed Successfully"
			End IF
		
			
		
		Case "ExitEndNote"
		
			Call ExitEndnote()
			If Window("EndNote").Exist(2) = False Then
			Status = "PASS"
			Expected = "EndnOte Application Must close Successfully"
			Actual = "Endnote Application was closed Successfully"
			
			else
			Status = "FAIL"
			Expected = "EndnOte Application Must close Successfully"
			Actual = "Endnote Application was NOT closed Successfully"
			End If
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	
	End Select
	
	
	GetTestStepObject = Expected&" - "&Actual&" - "&Status
	
	
End Function