'_____________________________________________________________________________________________________________________
'Description:To launch Endnote application
'Created by:Vidya Roobini
'Created date:04-Mar-2016
'Modified by:
'Modified date:
'Input:N/A
'Output:N/A
'syntax:call launchApplication()

'_____________________________________________________________________________________________________________________
'Delcare and Initialize variables/Flags
WinDowExists=True
BtnExist=True
ComboExist=True
DialogExist=True
WinObjExist=True
Public listCount
Public winObjText
Public MenuStatus
'________________________________________________________________________________
Public Function launchApplication()
        
		'SystemUtil.Run "C:\Program Files (x86)\EndNote X7\EndNote.exe"
		SystemUtil.Run strAppUrl
		wait(6)
End Function

'_____________________________________________________________________________________________________________________
'Description:To access gmail with valid credentials
'Created by:Vidya Roobini
'Created date:09-Mar-2016
'Modified by:
'Modified date:
'Input:Userid,Pwd
'Output:N/A
'syntax:call LoginGmail("abc@gmail.com","Sample")

'_____________________________________________________________________________________________________________________

Public Function LoginGmail(Userid,Pwd)

Set IE = CreateObject("InternetExplorer.Application")
IE.Visible = True
IE.Navigate "http://www.gmail.com"


Browser("Title:=Gmail.*").Page("title:=Gmail.*").Sync


Browser("Title:=Gmail.*").Page("title:=Gmail.*").WebEdit("name:=Email").Set Userid
Browser("Title:=Gmail.*").Page("title:=Gmail.*").WebButton("name:=Next").Click
Browser("Title:=Gmail.*").Page("title:=Gmail.*").WebEdit("name:=Passwd").Set Pwd
Browser("Title:=Gmail.*").Page("title:=Gmail.*").WebButton("name:=Sign in").click 
wait(5)

If Browser("name:=Inbox.*").Page("title:=Inbox.*").Exist(12) Then
  reporter.ReportEvent micPass ,"user "&Userid& " is logged in","User has succcessfully logged in "
else
   reporter.ReportEvent micFail,"User has not logged in ","user "&User2& "is logged in" 
End if

End Function

'_____________________________________________________________________________________________________________________
'Description:To close endnote application
'Created by:Vidya Roobini
'Created date:09-Mar-2016
'Modified by:
'Modified date:
'Input:N/A
'Output:N/A
'syntax:call ExitEndnote()

'_____________________________________________________________________________________________________________________

Public Function ExitEndnote()
	If Window("EndNote").Exist(2) Then
		Window("EndNote").Activate
		Window("EndNote").WinMenu("Menu").Select "File;Exit	Ctrl+Q"
    Else
	    window("EndNote").Activate
	    window("EndNote").WinMenu("Menu").Select "File;Exit	Ctrl+Q"
    End If
    wait(12)
End Function


'_____________________________________________________________________________________________________________________
'Description:To open a word document
'Created by:Vidya Roobini
'Created date:09-Mar-2016
'Modified by:
'Modified date:
'Input:document-path of document to open
'Output:N/A
'syntax:call OpenWordDoc("D:\sample.doc")

'_____________________________________________________________________________________________________________________
Public Function OpenWordDoc(document)
	set Word1=createobject("word.application")
    set doc1=word1.Documents.open(document)
    word1.Visible=true
End Function
'____________________________________________________________________________
'Function to CLick on Menu
Public Function ClickMenu(datMnuOption)

	Window("text:="&oEndNote_Win&".*").WinMenu("menuobjtype:=2").Select datMnuOption
	
End Function
'____________________________________________________________________________
'Function to Close EndNote
Public Function CloseEndNote()

	Window("text:="&oEndNote_Win&".*").Close
	
End Function
'____________________________________________________________________________
'Function for Sync
Public Function SynchWait(datWait)
	
	wait (datWait)
	
End Function
'____________________________________________________________________________
'Function to Open Sample Application if not already opened
Public Function OpenSampleLilbrary()
	
	If NOT Window("EndNote").Window("Sample_Library_X7").Exist(2) Then
	
		Window("EndNote").WinMenu("Menu").Select "File;Open Library...	Ctrl+O"
		
		If Window("EndNote").Dialog("Select a Reference Library:").Exist then 
			Window("EndNote").Dialog("Select a Reference Library:").WinEdit("File name:").Set "C:\Users\Public\Documents\EndNote\Examples\Sample_Library_X7.enl"
			'Window("EndNote").Dialog("Select a Reference Library:").WinObject("Items View").WinList("Items View").Select DataTable("Lib_Path",dtglobalsheet)
			wait 0.5
		    Window("EndNote").Dialog("Select a Reference Library:").WinObject("Open").Click
		End If
		
		'Window("EndNote").Window("Sample_Library_X7").Maximize
	
	End If
	
	Window("EndNote").Window("Sample_Library_X7").Maximize

End Function
'____________________________________________________________________________________________________________________________________
'Function to use the special keys.For ex: 'SHIFT' + 'F4'.Its a WSH function
Function TypeFunctionKeyW(Splkey, InputKey)

	Set wsh = CreateObject("WScript.Shell")

	wsh.SendKeys "" & Splkey & "{" & InputKey & "}"     
    
	Set wsh = Nothing
		
	Wait(2)
		
End Function
''____________________________________________________________________________________________________________________________________
'Function to rename folder
Public Function RenameFolder(datSrc,datDest)
	
	Set objFSO = CreateObject("Scripting.FileSystemObject")
	objFSO.MoveFolder datSrc,datDest
	
Set fso = Nothing
	
End Function
''____________________________________________________________________________________________________________________________________

'Open Normal LibraryFile 
Function OpenNormalLilbraryFile(LibraryFileName)
	
		Window("EndNote").WinMenu("Menu").Select "File;Open Library...	Ctrl+O"
		
		If Window("EndNote").Dialog("Select a Reference Library:").Exist then 
			Window("EndNote").Dialog("Select a Reference Library:").WinEdit("File name:").type "C:\Users\Public\Documents\EndNote\Examples\"&LibraryFileName
			'Window("EndNote").Dialog("Select a Reference Library:").WinObject("Items View").WinList("Items View").Select DataTable("Lib_Path",dtglobalsheet)
		    Window("EndNote").Dialog("Select a Reference Library:").WinObject("Open").Click
		End If
		
		'Window("EndNote").Window("Sample_Library_X7").Maximize
	Window("EndNote").Window("text:="&LibraryFileName).Maximize

End Function
''____________________________________________________________________________________________________________________________________
'Load repository Dynamically

Function LoadRepository(RepoPath)
'		Dim Repo
'		Set Repo = CreateObject("QuickTest.Application") 
'		Dim qtpRepositories 
'		Set qtpRepositories = Repo.Test.Actions("Main").ObjectRepositories 
'		qtpRepositories.Add("D:\EndNoteTesting\Repository\shared_repository.tsr")

RepositoriesCollection.Add(RepoPath)
End Function
''____________________________________________________________________________________________________________________________________

'Function for FodlerOperation

Function DeleteLibraryFolder(FolderName)
Dim EndNoteFolderObject
Set EndNoteFolderObject = CreateObject("Scripting.filesystemObject")
On error resume next 
wait(1)
	EndNoteFolderObject.DeleteFolder(FolderName)
	'Reporter.ReportEvent micPass,"Expected: Library Folder must not be deleted","Actual: "&err.description
	err.clear		
End Function
''____________________________________________________________________________________________________________________________________
'Delete Library File
Function DeleteLibraryFile(FolderName)
Dim EndNoteFolderObject, Result
Set EndNoteFolderObject = CreateObject("Scripting.filesystemObject")
On error resume next 
	
	result = EndNoteFolderObject.DeleteFile(FolderName)
	DeleteLibraryFile = result
	
	err.clear		
End Function
''____________________________________________________________________________________________________________________________________
'Library folder Exist
Function FolderExist(FolderName)
Dim EndNoteFolderObject
Set EndNoteFolderObject = CreateObject("Scripting.filesystemObject")
Dim result
	result = EndNoteFolderObject.FolderExists(FolderName)
	FolderExist = result
	
	Set EndNoteFolderObject = Nothing
End Function
''____________________________________________________________________________________________________________________________________
'Library file Exist
Function FileExist(FolderName)
Dim EndNoteFolderObject
Set EndNoteFolderObject = CreateObject("Scripting.filesystemObject")
Dim result
	result = EndNoteFolderObject.FileExists(FolderName)
	fileExist = result
	
End Function
''____________________________________________________________________________________________________________________________________

'CreateNew Library
Function CreateNewLibrary(LibraryName)
		Window("EndNote").WinMenu("Menu").Select "File;New..."
		Window("EndNote").Dialog("New Reference Library").WinEdit("File name:").Type "C:\Users\Public\Documents\EndNote\Examples\"&LibraryName
		Window("EndNote").Dialog("New Reference Library").WinButton("Save").Click
		Window("EndNote").Window("text:="&LibraryName).Maximize

End Function
'________________________________________________________________________________________________________________________________________________________________________________________

'Create New Reference
Function CreateNewReference(ReferenceName)
		Window("EndNote").WinMenu("Menu").Select "References;New Reference	Ctrl+N"
		Window("EndNote").Window("New Reference").WinObject("Author").Type ReferenceName
		Window("EndNote").Window("New Reference").WinObject("Issue").Type ReferenceName
		Window("EndNote").Window("New Reference").WinObject("Journal").Type ReferenceName	
		Window("EndNote").Window("New Reference").WinButton("FileAttachement").Click
		Window("EndNote").Dialog("SelectAFileToLink").Activate
		Window("EndNote").Dialog("SelectAFileToLink").WinEdit("FileName").Type "C:\Program Files (x86)\EndNote X7\EndNoteX7WinHelp.pdf"
		Window("EndNote").Dialog("SelectAFileToLink").WinButton("Open").Click
End Function
'________________________________________________________________________________________________________________________________________________________________________________________
'Email attached compressed library to the Recepient
Function EmailLibraryFile(EmailAddress)
							Window("EndNote").Window("EndNoteEmail").WinObject("RichEdit20WPT").Type EmailAddress
							Window("EndNote").Window("EndNoteEmail").WinObject("Subject").Type FileName
							Window("EndNote").Window("EndNoteEmail").WinObject("Message").Type FileName
							Window("EndNote").Window("EndNoteEmail").WinButton("Send").Click
		
End Function
'________________________________________________________________________________________________________________________________________________________________________________________
'Verify if Window Exists
Public Function Verify_Window_Exists(WinText)

	If Window("text:="&oEndNote_Win&".*").Window("text:="&WinText&"").Exist Then
	
		WinDowExists=True
		
	Else
	
		WinDowExists=False

	End If
	
	
End Function
'_____________________________________________________________________________________________________________________________________________________________________________________
'Function to Create and Add ObjectData
Public Function AddObjectData(strConNo,strSheetNm)
		
	DataTable.AddSheet (strSheetNm)
	
	'	DataTable.ImportSheet 
	DataTable.ImportSheet  strObjRepoPath, strConNo, strSheetNm
	WinRow=DataTable.GetSheet(strSheetNm).GetRowCount
		
	filetxt.WriteLine ("'***************************"&strSheetNm&"****************************************")
	'filetxt.WriteLine (datatable.GetSheet("Window").GetParameter("Variable").ValueByRow(2)&"="&datatable.GetSheet("Window").GetParameter("Value_001").ValueByRow(2))
	
	For i = 1 To WinRow
		
		filetxt.WriteLine (datatable.GetSheet(strSheetNm).GetParameter("Variable").ValueByRow(i)&"="""&datatable.GetSheet(strSheetNm).GetParameter("Value_001").ValueByRow(i))&""""
		filetxt.WriteLine (datatable.GetSheet(strSheetNm).GetParameter("Variable").ValueByRow(i)&"2"&"="""&datatable.GetSheet(strSheetNm).GetParameter("Value_002").ValueByRow(i))&""""
	'filetxt.WriteLine "jay"
	
	Next
	filetxt.WriteLine ("'________________________________________________________________________")
		
End Function
'________________________________________________________________________________________________________________________________________________________________________________________
'Function to Click on the VinTreeView link
Public Function Click_WinTreeView_Link(WinId,WinFrame,WinFrame2,WinTreeLink)

	Window("text:="&oEndNote_Win&".*").Window("window id:="&WinId&"").WinTreeView("window id:="&WinFrame&"","nativeclass:="&WinFrame2&"").Select WinTreeLink
	
End Function
'________________________________________________________________________________________________________________________________________________________________________________________
'Function to Close any Non Parent window:
Public Function CloseWindow(WinId)

	Window("text:="&oEndNote_Win&".*").Window("window id:="&WinId&"").close
	
End Function
'________________________________________________________________________________________________________________________________________________________________________________________
'Function to ClickButton:Jay:8-12-16
Public Function ClickButton(WinId,ButtonName)

	Window("text:="&oEndNote_Win&".*").Window("window id:="&WinId&"").WinButton("text:="&ButtonName&"").Click
	
End Function
'________________________________________________________________________________________________________________________________________________________________________________________
'Function to Verify if Button exists :Jay:8-12-16
Public Function VerifyButtonExists(WinId,ButtonName)

	If Window("text:="&oEndNote_Win&".*").Window("window id:="&WinId&"").WinButton("text:="&ButtonName&"").Exist Then
	
		BtnExist=True
		
	Else
	
		BtnExist=False
		
	End If
End Function
'________________________________________________________________________________________________________________________________________________________________________________________
'Function to Select an option from Combobox:Jay:8-12-16
Public Function SelectComboboxOption(WinId,ComboboxName,ComboboxName2,ComboOption)

	Window("text:="&oEndNote_Win&".*").Window("window id:="&WinId&"").WinComboBox("attached text:="&ComboboxName&"","nativeclass:="&ComboboxName2&"").Select ComboOption
	
End Function
'________________________________________________________________________________________________________________________________________________________________________________________
'Function to Type is WinObject field-Jay:8-12-16
Public Function TypeInWinObject(WinId,WinObjName,WinObjData)

	Window("text:="&oEndNote_Win&".*").Window("window id:="&WinId&"").WinObject("window id:="&WinObjName&"").Type WinObjData
	
End Function
'________________________________________________________________________________________________________________________________________________________________________________________
'Function to Verify if Combobox exists :Jay:8-12-16
Public Function VerifyComboboxExists(WinId,ComboboxName,ComboboxName2)

	If Window("text:="&oEndNote_Win&".*").Window("window id:="&WinId&"").WinComboBox("attached text:="&ComboboxName&"","nativeclass:="&ComboboxName2&"").Exist Then
	
		ComboExist=True
		
	Else
	
		ComboExist=False
		
	End If
End Function
'________________________________________________________________________________________________________________________________________________________________________________________
'Verify if New Window Exists 
Public Function VerifyNewWindowExists(WinId)

	If Window("text:="&oEndNote_Win&".*").Window("window id:="&WinId&"").Exist Then
	
		WinDowExists=True
		
	Else
	
		WinDowExists=False

	End If
	
End Function
'_____________________________________________________________________________________________________________________________________________________________________________________
'Verify Window text by Getting RO property
Public Function VerifyWindowText(WinId,WinText)

	Call SynchWait(2)

	windowText=Window("text:="&oEndNote_Win&".*").Window("window id:="&WinId&"").GetROProperty("text")
	
	If InStr(windowText,WinText)>0 Then
		
		WinDowExists=True
		
	Else
	
		WinDowExists=False

	End If
	
End Function
'_____________________________________________________________________________________________________________________________________________________________________________________
'Verify if Dialog exists-Jay:9-12-16
Public Function VerifyDialogExists(DialogTxt)

	If Window("text:="&oEndNote_Win&".*").Dialog("text:="&DialogTxt&"").Exist then
	
		DialogExist=True
		
	Else
	
		DialogExist=False

	End If
		
End Function
'_____________________________________________________________________________________________________________________________________________________________________________________
'Click on Button within the Dialog-Jay:9-12-16
Public Function ClickDialogButton(DialogTxt,ButtonText)

	'Window("text:="&oEndNote_Win&".*").Dialog("text:="&DialogTxt&"").WinButton("text:="&ButtonText&"").Click
	Window("EndNote").Dialog("text:="&DialogTxt&"").WinButton("text:="&ButtonText&"").Click
	
End Function
'_____________________________________________________________________________________________________________________________________________________________________________________
'Type Data in the WinObject within the Dialog-Jay:9-12-16
Public Function TypeDialogWinObject(DialogTxt,WinObjText,strTestData)

	Window("text:="&oEndNote_Win&".*").Dialog("text:="&DialogTxt&"").WinObject("attachedtext:="&WinObjText&"").Type strTestData
	
End Function
'_____________________________________________________________________________________________________________________________________________________________________________________
'Verify if Dialog exists-Jay:9-12-16
Public Function VerifyDialogWithinDialogExists(DialogTxt,ChildDialogTxt)

	'If Window("text:="&oEndNote_Win&".*").Dialog("text:="&DialogTxt&"").Dialog("text:="&ChildDialogTxt&"").Exist then
	If Window("EndNote").Dialog("text:="&DialogTxt&"").Dialog("text:="&ChildDialogTxt&"").Exist then
	
		DialogExist=True
		
	Else
	
		DialogExist=False

	End If
	
			
End Function
'_____________________________________________________________________________________________________________________________________________________________________________________
'ClickDialogButton within the Dialog-Jay:9-12-16
Public Function ClickDialogButtonWithinDialog(DialogTxt,ChildDialogTxt,WinButton)


	'Window("text:="&oEndNote_Win&".*").Dialog("text:="&DialogTxt&"").Dialog("text:="&ChildDialogTxt&"").WinButton("text:="&WinButton&"").Click
	Window("EndNote").Dialog("text:="&DialogTxt&"").Dialog("text:="&ChildDialogTxt&"").WinButton("text:="&WinButton&"").Click
	
End Function
'_____________________________________________________________________________________________________________________________________________________________________________________
'Type Data in the Dialog WinEdit within the Dialog-Jay:9-12-16
Public Function TypeDialogWinEditWithinDialog(DialogTxt,ChildDialogTxt,WinEditTxt,strTestData)

	Window("text:="&oEndNote_Win&".*").Dialog("text:="&DialogTxt&"").Dialog("text:="&ChildDialogTxt&"").WinEdit("attached text:="&WinEditTxt&"","regexpwndclass:=Edit").Type strTestData
	
End Function
'_____________________________________________________________________________________________________________________________________________________________________________________
'Click Winobject in the Dialog WinObject within the Dialog-Jay:9-12-16
Public Function ClickDialogWinObjectWithinDialog(DialogTxt,ChildDialogTxt,WinObjTxt)

	Window("text:="&oEndNote_Win&".*").Dialog("text:="&DialogTxt&"").Dialog("text:="&ChildDialogTxt&"").WinObject("text:="&WinObjTxt&"").Click
	
End Function
'_____________________________________________________________________________________________________________________________________________________________________________________
'Function to Select an option from Combobox in Dialog:Jay:9-12-16
Public Function SelectDialogComboboxOption(DialogTxt,ComboboxName,ComboboxName2,ComboOption)

	Window("text:="&oEndNote_Win&".*").Dialog("text:="&DialogTxt&"").WinComboBox("window id:="&ComboboxName&"","nativeclass:="&ComboboxName2&"").Select ComboOption
	
End Function
'________________________________________________________________________________________________________________________________________________________________________________________
'Get WinslitView Item in a Dialog within dialog-Jay:9-12-16
Public Function GetSubItemDialogWinlistViewWithinDialog(DialogTxt,ChildDialogTxt,WinLIstView,strTestData)

	Window("text:="&oEndNote_Win&".*").Dialog("text:="&DialogTxt&"").Dialog("text:="&ChildDialogTxt&"").WinListView("window id:="&WinLIstView&"").GetSubItem strTestData
	
End Function
'________________________________________________________________________________________________________________________________________________________________________________________
'Get WinslitView SubItem-Jay:9-12-16
Public Function GetItemCountWinlistView(WinTxt,WinListView)

	listCount=Window("text:="&oEndNote_Win&".*").Window("text:="&WinTxt&"").WinListView("window id:="&WinListView&"").GetItemsCount()
	
	
End Function
'________________________________________________________________________________________________________________________________________________________________________________________
'Select WinslitView SubItem-Jay:9-12-16
Public Function SelectWinlistViewItem(WinTxt,WinListView,strIndex,MouseBtn)

	Window("text:="&oEndNote_Win&".*").Window("text:="&WinTxt&"").WinListView("window id:="&WinListView&"").Select strIndex,MouseBtn,100	
		
End Function
'________________________________________________________________________________________________________________________________________________________________________________________
'Function to Cick on Context Menu-Jay:9-12-16
Public Function ClickContextMenu(MnuOption)

	Window("text:="&oEndNote_Win&".*").WinMenu("menuobjtype:=3").Select MnuOption
	
End Function
'____________________________________________________________________________
'Get Text from Winobject within new Window-Jay:9-12-16
Public Function GetWinObjectText(WinId)

	winObjText=Window("text:="&oEndNote_Win&".*").WinObject("window id:="&WinId&"","nativeclass:=R.*").GetROProperty("text")
	

End Function
'_____________________________________________________________________________________________________________________________________________________________________________________
'Function to Verify if WinObject exists :Jay:20-12-16
Public Function VerifyWinObjectExists(WinId,WinObj)

	If Window("text:="&oEndNote_Win&".*").Window("window id:="&WinId&"").WinObject("window id:="&WinObj&"").Exist Then
	
		WinObjExist=True
		
	Else
	
		WinObjExist=False
		
	End If
End Function
'________________________________________________________________________________________________________________________________________________________________________________________
'Function to get ROP Property of Winobject-Jay:20-12-16
Public Function WinObjectEnabled(WinId,WinObjName,ROProp)

	WinObjEnabled=Window("text:="&oEndNote_Win&".*").Window("window id:="&WinId&"").WinObject("window id:="&WinObjName&"").GetROProperty(ROProp)
		
End Function
'________________________________________________________________________________________________________________________________________________________________________________________
Public Function CloseLibraryFile()

		Window("EndNote").WinMenu("Menu").Select "File;Close Library	Ctrl+W"	
	
End Function
'_______________________________________________________________________________________________________________________________________________________________________
Public Function CloseReferenceFile()
	
	Window("EndNote").WinMenu("Menu").Select "File;Close Reference	Ctrl+W"
	
End Function
'_______________________________________________________________________________________________________________________________________________________
'Function to Set WinRadioButton-Jay:21-12-16
Public Function SetWinRadioButton(WinId,WinRadioBtnTxt)

	Window("text:="&oEndNote_Win&".*").Window("window id:="&WinId&"").WinRadioButton("text:="&WinRadioBtnTxt&"").Set()
		
End Function
'________________________________________________________________________________________________________________________________________________________________________________________
'Function to Select an option from Combobox identified by window id:Jay:21-12-16
Public Function SelectWinIdComboboxOption(WinId,ComboboxWinId,ComboboxWinId2,ComboOption)

	Window("text:="&oEndNote_Win&".*").Window("window id:="&WinId&"").WinComboBox("window id:="&ComboboxName&"","nativeclass:="&ComboboxName2&"").Select ComboOption
	
End Function
'________________________________________________________________________________________________________________________________________________________________________________________
'Function to get Options count available in Combobox-Jay:21-12-16
Public Function GetComboboxOptionsCount(WinId,ComboboxWinId,ComboboxWinId2)

	comboxItemsCount=WinObjEnabled=Window("text:="&oEndNote_Win&".*").Window("window id:="&WinId&"").WinComboBox("window id:="&ComboboxName&"","nativeclass:="&ComboboxName2&"").GetItemsCount ()
		
End Function
'________________________________________________________________________________________________________________________________________________________________________________________