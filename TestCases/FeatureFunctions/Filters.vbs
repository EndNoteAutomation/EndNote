'**************************************
'Functionality:Filters
'Purpose:Filters feature test cases
'Created by:Jayateerth
'Created Date:1/12/2016
'Modified by;
'Modified date:
'**************************************
Option Explicit 
Dim path,TesetStepExecuteName,oShell,TestDataArray,TestCaseName,strNewFilterMenu,WinDowExists,strTextFile,strLine1,strLine2,strData1,filesysO,filetxtO,strFilter,strLine3,strLine4,strLine5,strLine6,DialogExist,WinObjEnabled,comboxItemsCount
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
			
		Case "CreaterTextFile"
		
			TestDataArray=Split(TestDataArray,"|")
				'Get the Author Name
			strTextFile=TestDataArray(0)
			strLine1=TestDataArray(1)
			strLine2=TestDataArray(2)
		
			'Create the Text file
			Set filesysO = CreateObject("Scripting.FileSystemObject") 
			Set filetxtO = filesysO.CreateTextFile(strTextFile,True) 
					
			path = filesysO.GetAbsolutePathName("D:\Sample.txt") 
			getname = filesysO.GetFileName(path) 
			filetxtO.WriteLine(strLine1) 
			filetxtO.WriteLine(strLine2) 
			
			filetxtO.Close 		
			
			Status = "PASS"
			Expected = "Text File to be imported later is created."
		
			Actual = "Text File to be imported later is created."
			
		Case "OpenSampleLibrary"
			
			'Open Sample LIb
			Call OpenSampleLilbrary()
			
			
			'Verify
			Call Verify_Window_Exists(oSampleLibrary)
									
			If WinDowExists=True then
			
				Status = "PASS"
				Expected = "Sample Library Must be Launched"
				Actual = "Sample Library is Launched"
			Else
				Status = "FAIL"
				Expected = "Sample Library Must	 be Launched"
				Actual = "Sample Library is Not Launched"
			
			End If
			
		Case "OpenUntitledFilter"
			
			'ClickMenu
			TestDataArray=Split(TestDataArray,"|")
			strNewFilterMenu=TestDataArray(0)
			
			Call ClickMenu(strNewFilterMenu)
			
			Call VerifyWindowText(oDynamicWindow_Win,oUnTitltedFilter_Win)
												
			If WinDowExists=True then
			
				Status = "PASS"
				Expected = "Open Untitled Filter window.Untitled window must be opened."
				Actual = "Untitled window is opened."
			Else
				Status = "FAIL"
				Expected = "Open Untitled Filter window.Untitled window must be opened."
				Actual = "Untitled window is not opened."
			
			End If
			
						
		Case "ClickOnTempletes"
		
			'Select Templates from Left frame
			TestDataArray=Split(TestDataArray,"|")
			strData1=TestDataArray(0)
						
			Call Click_WinTreeView_Link(oDynamicWindow_Win,oWinTreeView_WinTree,oWinTreeView_WinTree2,strData1)
				
			'Verify if Templates section is displyed
			Call VerifyButtonExists(oDynamicWindow_Win,oInsertField_Button)
			
			If BtnExist=True Then
			
				Status = "PASS"
				Expected = "Select 'Templates' from left window.Templates information is displayed"
				Actual = "Templates information is displayed"
			Else
				Status = "FAIL"
				Expected = "Select 'Templates' from left window..Templates information is displayed."
				Actual = "Templates information is not displayed"
			
			End if
		
		Case "SelectGeneric"
		
			'Select Generic option from Reference Types combo box
			TestDataArray=Split(TestDataArray,"|")
			strData1=TestDataArray(0)
			
			Call SelectComboboxOption(oDynamicWindow_Win,oReferenceTypes_ComboBox,oNameOrder_ComboBox2,strData1)
		
			Status = "PASS"
			Expected = "Select 'Generic' from Reference Type Combobox.'Generic' is selected from Reference Type Combobox."
			Actual = "'Generic' is selected from Reference Type Combobox."
		
		
		Case "TypeTagData"
		
			'Type Tag data
			TestDataArray=Split(TestDataArray,"|")
			strData1=TestDataArray(0)
			
			Call TypeInWinObject(oDynamicWindow_Win,oTag_WinObject,strData1)
			Status = "PASS"
			Expected = "Type 'AU' in Tag cell.Data is entered in Tag cell."
			Actual = "Data is entered in Tag cell."
				
		Case "TypeFieldsData"
			
			'Type Fields data
			TestDataArray=Split(TestDataArray,"|")
			strData1=TestDataArray(0)
			
			Window("EndNote").InsightObject("Feilds_Cell").Click
			'Call TypeInWinObject(oDynamicWindow_Win,oFields_WinObject,strData1)
			Window("EndNote").InsightObject("Feilds_Cell").Type strData1
			Status = "PASS"
			Expected = "Type 'Author' in Fields cell.Data is entered in Fields cell."
			Actual = "Data is entered in Fields cell."
			
					
		Case "ClickOnAuthorParsing"
			'Click on Author Parsing left menu
			TestDataArray=Split(TestDataArray,"|")
			strData1=TestDataArray(0)
			
			Call Click_WinTreeView_Link(oDynamicWindow_Win,oWinTreeView_WinTree,oWinTreeView_WinTree2,strData1)
				
			'Verify if Author Parsing window is opened
			Call VerifyComboboxExists(oDynamicWindow_Win,oNameOrder_ComboBox,oNameOrder_ComboBox2)
			If ComboExist=True Then
			
				Status = "PASS"
				Expected = "Select 'Author Parsing from left menu.Author Parsing information is displayed"
				Actual = "Author Parsing information is displayed"
			Else
				Status = "FAIL"
				Expected = "Select 'Author Parsing from left menu.Author Parsing information is displayed."
				Actual = "Author Parsing information is not displayed"
			
			End if
		
		
		Case "SelectFromNameOrder"
			'Select Last Always Preceeds First from Name Order combobox
			TestDataArray=Split(TestDataArray,"|")
			strData1=TestDataArray(0)	
			Call SelectComboboxOption(oDynamicWindow_Win,oNameOrder_ComboBox,oNameOrder_ComboBox2,strData1)
		
			Status = "PASS"
			Expected = "Select "&strData1&"."&strData1&" is selected from Name Order Combobox."
			Actual = ""&strData1&"is selected from Name Order Combobox."
		
		Case "CloseNewFilterWindow"
		
			'Close New Filter Window
			Call CloseWindow(oDynamicWindow_Win)
			
			
			
			'Verify if Confirmation dialog is displayed
			Call VerifyDialogExists(oEndNote_Dialog)		
						
			If DialogExist=True Then
		
				Status = "PASS"
				Expected = "Close New Filter window.Confirmation Dialog is displayed."
				Actual = "Confirmation Dialog is displayed."
			Else
				Status = "FAIL"
				Expected = "Confirmation Dialogbox is displayed."
				Actual = "Close New Filter window..Confirmation Dialogbox is not displayed."
			
			End If
		
		Case "ClickYesButton"
			'Click on Yes Dialog button
			Call ClickDialogButton(oEndNote_Dialog,oYes_Button)
				
			'Verify if Save As dialog is displayed
			Call VerifyDialogExists(oSaveAs_Dialog)		
			
			If DialogExist=True Then
		
				Status = "PASS"
				Expected = "Click on Yes button.Save As Dialog is displayed."
				Actual = "Save As Dialog is displayed."
			Else
				Status = "FAIL"
				Expected = "Save As Dialog is displayed."
				Actual = "Click on Yes button.Save As Dialog is not displayed."
			
			End If
		
		Case "TypeNewFilterName"
			'Type Filter Name in the Dialog field
			TestDataArray=Split(TestDataArray,"|")
			strFilter=TestDataArray(0)
			Call TypeDialogWinObject(oSaveAs_Dialog,oFilterName_WinObject,strFilter)		
			
			Status = "PASS"
			Expected = "Type New Fitler Name.Filter Name is entered."
			Actual = "Filter Name is entered."	
				
		Case "ClickSaveButton"
			'Click on the Save button
			Call ClickDialogButton(oSaveAs_Dialog,oSave_Button)
			
			'Confirm replace FIlter name message
			Call VerifyDialogWithinDialogExists(oSaveAs_Dialog,oEndNote_Dialog)
			
			If DialogExist=True Then 
			
				Call ClickDialogButtonWithinDialog(oSaveAs_Dialog,oEndNote_Dialog,oYes_Button)
			
			End if
		
		
			'Verify if Sample Library window is displayed
			Call VerifyNewWindowExists(oDynamicWindow_Win)
			
			If WinDowExists=False then
			
				Status = "PASS"
				Expected = "Click Save button.Untitled Filter window is closed."
				Actual = "Untitled Filter window is closed."				

			Else
			
				Status = "FALSE"
				Expected = "Click Save button.Untitled Filter window is closed."
				Actual = "Untitled Filter window is not closed."
			
			End if		
		
		Case "ClickImportFileMenu"
							
			Call SynchWait(1)
			'Click menu-File;Import;File...
			TestDataArray=Split(TestDataArray,"|")
			strData1=TestDataArray(0)
			
			Call ClickMenu(strData1)
			
			'Verify if Import File Dialog displayed
			Call VerifyDialogExists(oImportFile_Dialog)		
			
			If DialogExist=True Then
		
				Status = "PASS"
				Expected = "Click on File;Import;File...menu.Import File Dialog is displayed."
				Actual = "Import File Dialog is displayed."
			Else
				Status = "FAIL"
				Expected = "Click on File;Import;File...menu.Import File Dialog is displayed."
				Actual = "Import File Dialog is not displayed."
			End If
				
		Case "ClickChooseDotButton"
			'Click on Choose... button
			Call ClickDialogButton(oImportFile_Dialog,oChooseDot_Button)
			
			'Verify if Open Dialog is displayed
			Call VerifyDialogWithinDialogExists(oImportFile_Dialog,oOpen_Dialog)
			
			If DialogExist=True Then
		
				Status = "PASS"
				Expected = "Click Choose...button.Open Dialog is displayed."
				Actual = "Open Dialog is displayed."
			Else
				Status = "FAIL"
				Expected = "Click Choose...button.Open Dialog is displayed."
				Actual = "Open Dialog is not displayed."
			
			End If
		
		
		Case "TypeFileName"
		
			'Type File name created above
			Call TypeDialogWinEditWithinDialog(oImportFile_Dialog,oOpen_Dialog,oFileName_WinEdit,strTextFile)
		
			Status = "PASS"
			Expected = "File Name created above is entered."
			Actual = "File Name created above is entered."
		
		Case "ClickOpenButton"
			'Click on Open button in Dialog
			Call ClickDialogWinObjectWithinDialog(oImportFile_Dialog,oOpen_Dialog,oOpenBtn_WinObject)	
			Status = "PASS"
			Expected = "Open Button is clicked."
			Actual = "Open Button is clicked."		
		
		Case "SelectOtherFilters"
			'Select Other Filters from Import Options Combobox		
			TestDataArray=Split(TestDataArray,"|")
			strData1=TestDataArray(0)
			
			Call SelectDialogComboboxOption(oImportFile_Dialog,oImportOption_ComboBox,oImportOption_ComboBox2,strData1)
			
						
			'Verify if Choose An Option dialog is opened
			Call VerifyDialogWithinDialogExists(oImportFile_Dialog,oChooseAnImportFile_Dialog)
			
			If DialogExist=True Then
		
				Status = "PASS"
				Expected = "Select Other Filters...from dropdown.Choose An Import Dialog is displayed."
				Actual = "Choose An Import Dialog is displayed."
			Else
				Status = "FAIL"
				Expected = "Select Other Filters...from dropdown.Choose An Import Dialog is displayed."
				Actual = "Choose An Import Dialog is not displayed."
			
			End If
		
		
		Case "SelectCreatedFilter"
			'Select Previously created filter from WinlistView
			Call GetSubItemDialogWinlistViewWithinDialog(oImportFile_Dialog,oChooseAnImportFile_Dialog,oFilter_WinlistView,strFilter)
			
			Status = "PASS"
			Expected = "Select Filter created earlier.Previously created Filter is selected from WinlistView."
			Actual = "Previously created Filter is selected from WinlistView."
				
		Case "ClickChooseButton"
			'Click on Choose button 
			Call ClickDialogButtonWithinDialog(oImportFile_Dialog,oChooseAnImportFile_Dialog,oChoose_Button)
			
			Status = "PASS"
			Expected = "Click Choose button.Choose button is clicked."
			Actual = "Choose button is clicked."
		
		Case "CickImportButton"
			'Click on Import button
			Call ClickDialogButton(oImportFile_Dialog,oImport_Button)

			'Verify if Reference is imported
			Call GetItemCountWinlistView(oSampleLibrary,oReferences_WinlistView)
		
			If listCount>0 Then
			
				Status = "PASS"
				Expected = "Click Import button.Reference is imported."
				Actual = "Reference is imported."
					
			Else
				Status = "FAIL"
				Expected = "Click Import button.Reference is imported."
				Actual = "Reference is not imported."		
			
			End If
		
		Case "SelectNewlyImportedReference"
			'Select newly Imported reference
			Call SelectWinlistViewItem(oSampleLibrary,oReferences_WinlistView,0,micLeftBtn)
			Status = "PASS"
			Expected = "Select Newly imported Reference.Newly imported Reference is selected."
			Actual = "Newly imported Reference is selected."		
		
		Case "ClickEditReferencesMenu"
			'Click End References from Context Menu
			TestDataArray=Split(TestDataArray,"|")
			strData1=TestDataArray(0)
			'Call SelectWinlistViewItem(oSampleLibrary,oReferences_WinlistView,0,micLeftBtn)
			Call SelectWinlistViewItem(oSampleLibrary,oReferences_WinlistView,0,micRightBtn)
						
			'Call ClickMenu(strData1)
			Call ClickContextMenu(strData1)
		
			Status = "PASS"
			Expected = "Edit new Reference.Edit Reference Option is selected."
			Actual = "Edit new Reference.Edit Reference Option is selected."		
		
		Case "VerifyAuthorDataExists"
			'Verify if Author Data exists
			Call GetWinObjectText(oAuthor_WinObject)			
									
			If winObjText<>"" Then
			
				Status = "PASS"
				Expected = "Verify Author data exists.Author Names displayed."
				Actual = "Author Names displayed."
			Else
				Status = "FAIL"
				Expected = "Verify Author data exists.Author Names displayed."
				Actual = "Author Names not displayed."
			
			End If
		
		Case "CloseEndNote"
			'Close EndNote
			Call CloseEndNote()
			Status = "PASS"
			Expected = "Close EndNote.EndNote is closed."
			Actual = "EndNote is closed."	
		
		Case "SelectFromInterpretFirstNameAs"
			'Select from Name Interpret First Name As combobox
			TestDataArray=Split(TestDataArray,"|")
			strData1=TestDataArray(0)	
			Call SelectComboboxOption(oDynamicWindow_Win,oInterpretFirstNameAs_Combobox,oInterpretFirstNameAs_Combobox2,strData1)
		
			Status = "PASS"
			Expected = "Select "&strData1&"."&strData1&" is selected from Interpret First Name As Combobox."
			Actual = ""&strData1&"is selected from Interpret First Name As Combobox."
		
		Case "CloseNewWindow"
		
			Call CloseWindow(oDynamicWindow_Win)
			
			Status = "PASS"
			Expected = "Close New Filter window.New Filter Window is closed."
			Actual = "New Filter Window is closed."
				
		Case "ClickNoButton"
		
			Call ClickDialogButton(oEndNote_Dialog,oNo_Button)	
			
		Case "SelectFromBetweenAuthors"
			'Select from Between Authors combobox
			TestDataArray=Split(TestDataArray,"|")
			strData1=TestDataArray(0)	
			Call SelectComboboxOption(oDynamicWindow_Win,oBetweenAuthors_Combobox,oBetweenAuthors_Combobox2,strData1)
		
			Status = "PASS"
			Expected = "Select "&strData1&"."&strData1&" is selected from Interpret First Name As Combobox."
			Actual = ""&strData1&"is selected from Interpret First Name As Combobox."			
			
		Case "SelectFromLastAndFirstNames"
			'Select from Between Authors combobox
			TestDataArray=Split(TestDataArray,"|")
			strData1=TestDataArray(0)	
			Call SelectComboboxOption(oDynamicWindow_Win,oLastAndFitstNames_Combobox,oLastAndFitstNames_Combobox2,strData1)
		
			Status = "PASS"
			Expected = "Select "&strData1&"."&strData1&" is selected from Last And First Names Combobox."
			Actual = ""&strData1&"is selected from Last And First Names Combobox."				
		
		Case "CreaterTextFile6Lines"
		
			TestDataArray=Split(TestDataArray,"|")
				'Get the Author Name
			strTextFile=TestDataArray(0)
			strLine1=TestDataArray(1)
			strLine2=TestDataArray(2)
			strLine3=TestDataArray(3)
			strLine4=TestDataArray(4)
			strLine5=TestDataArray(5)
			strLine6=TestDataArray(6)
			
		
			'Create the Text file
			Set filesysO = CreateObject("Scripting.FileSystemObject") 
			Set filetxtO = filesysO.CreateTextFile(strTextFile,True) 
					
			path = filesysO.GetAbsolutePathName("D:\Sample.txt") 
			getname = filesysO.GetFileName(path) 
			filetxtO.WriteLine(strLine1) 
			filetxtO.WriteLine(strLine2) 
			filetxtO.WriteLine(strLine3) 
			filetxtO.WriteLine(strLine4) 
			filetxtO.WriteLine(strLine5) 
			filetxtO.WriteLine(strLine6) 
					
			
			filetxtO.Close 		
			
			Status = "PASS"
			Expected = "Text File to be imported later is created."
		
			Actual = "Text File to be imported later is created."
		
		Case "VerifyWinObjExists"
			'Select from Name Interpret First Name As combobox
			Call VerifyWinObjectExists(oDynamicWindow_Win,oLastAndFirstNames_WinObject)
			
			If WinObjExist=True Then
			
				Status = "PASS"
				Expected = "Edit field is displayed next to Drop down."
				Actual = "Edit field is displayed next to Drop down."
			
			Else
			
				Status = "FAIL"
				Expected = "Edit field is displayed next to Drop down."
				Actual = "Edit field is not displayed next to Drop down."
			
			End If 
			
		Case "VerifyWInObjEnabled"
		'Select from Name Interpret First Name As combobox
		Call WinObjectEnabled(oDynamicWindow_Win,oLastAndFirstNames_WinObject,"enabled")
		
		If WinObjEnabled="True" Then
		
			Status = "PASS"
			Expected = "Edit Field is enabled."
			Actual = "Edit Field is enabled."
		
		Else
		
			Status = "FAIL"
			Expected = "Edit Field is enabled."
			Actual = "Edit Field is notenabled."
		
		End If 
		
		Case "TypeInEdifFld"
			TestDataArray=Split(TestDataArray,"|")
			strData1=TestDataArray(0)
			
			Call TypeInWinObject(oDynamicWindow_Win,oLastAndFirstNames_WinObject,strData1)
			Status = "PASS"
			Expected = "Data typed in Edit field."
			Actual = "Data typed in Edit field."		
					
		Case "SetRadioButton"
		
			Call SetWinRadioButton(oDynamicWindow_Win,oDiscard_WinRadioButton)
			
			Status = "PASS"
			Expected = "Set "&oDiscard_WinRadioButton&" radio button."&oDiscard_WinRadioButton&" radio button is set."
			Actual = ""&oDiscard_WinRadioButton&" radio button is set."
		
		
	Case "VerifyComboboxOptions"
		'Verify the Options available in the COmboxbox
		Call GetComboboxOptionsCount(oDynamicWindow_Win,oImportInto_Combobox,oImportInto_Combobox2)
		If comboxItemsCount >0 Then
		
			Status = "PASS"
			Expected = "Verify the items in the dropdown.All the generic fields are displayed."
			Actual = "All the generic fields are displayed."
		
		Else
		
			Status = "FAIL"
			Expected = "Verify the items in the dropdown.All the generic fields are displayed."
			Actual = "All the generic fields are not displayed."
		
		End If 
	
	End Select
	
	GetTestStepObject = Expected&" - "&Actual&" - "&Status
	
	
End Function