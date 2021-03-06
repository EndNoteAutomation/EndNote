'**************************************
'functionality:Reference Type/Data
'purpose:default fields of Reference type
'Created by:Dilip Parashuram 
'created Date:29/08/2016
'modified by;
'Modified date:
'**************************************
Option Explicit 
Dim path,TesetStepExecuteName,oShell
Set oShell =  CreateObject("WScript.shell")
'Load Library Functions & repository 
'Execution the Test Objects




Function GetTestStepObject(TesetStepExecuteName)

Dim Status, Actual, Expected
	Select Case TesetStepExecuteName
		Case "EN_Launch_EndNote"
				
		If Window("EndNote").Exist(3) = False Then
			call launchApplication()
		End If
		If Window("EndNote").Exist(2) = true Then
			Status = "PASS"
			Expected = "Endnote Must be Launched"
			Actual = "Endnote is Launched"
			else
			Status = "FAIL"
			Expected = "Endnote Must be Launched"
			Actual = "Endnote is Not Launched"
			
		End If
		
		Case "EN_Open_Sample_Library_File"
			call OpenSampleLilbrary()
			If Window("EndNote").Exist(2) = true Then
			Status = "PASS"
			else
			Status = "FAIL"
			
			End If
		
		Case "EN_Close_Library_Files"
			Window("EndNote").Activate
			oShell.SendKeys "^+W"
			If Window("EndNote").Exist(2) = true Then
			Status = "PASS"
			Expected = "All the Opened Libraries must be closed Successfully"
			Actual = "All the opened Libraries are closed Successfully"
			else
			Status = "FAIL"
			
			End If
		Case "EN_Open_Preferences"
			ClickMenu("Edit;Preferences...")
			If Window("EndNote").Dialog("EndNote Preferences").Exist = true Then
			Status = "PASS"
			else
			Status = "FAIL"
			End IF
		
		Case "EN_Select_Reference"
		
			Window("EndNote").Dialog("EndNote Preferences").WinTreeView("PreferenceOption").Select "Reference Types"
			If Window("EndNote").Dialog("EndNote Preferences").WinComboBox("DefaultReferenceType").Exist(2) =  True Then
			Status = "PASS"
			else
			Status = "FAIL"
			End IF
		
		Case "EN_Select_ReferenceType"
			Window("EndNote").Dialog("EndNote Preferences").WinComboBox("DefaultReferenceType").Select "Aggregated Database"
			oShell.SendKeys "ENTER"
			
		
		Case "EN_Click_Button"
			Window("EndNote").Dialog("EndNote Preferences").WinButton("ModifyReferenceTypes").Click

		Case "En_Verify_ReferenceType"
		
		Case "EN_Close_Modify_Reference_Type"
		
		Case "EN_Close_Preferences"
		
		Case "EN_Exit_EndNote"
		
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