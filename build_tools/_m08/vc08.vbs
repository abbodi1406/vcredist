' Usage:
'
' cscript <name_of_file>.vbs
'
' script by dumpydooby
' modded by ricktendo64
' updated by abbodi1406
Option Explicit
Dim ws, installer, fs, db, view, record, x, sProperty
Set ws = WScript.CreateObject("WScript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")
Set installer = WScript.CreateObject("WindowsInstaller.Installer")
If WScript.Arguments.Count <> 0 Then
	For each x in WScript.Arguments
		ProcessMSI x
	Next
Else
	If fs.FileExists("vcredist.msi") Then ProcessMSI "vcredist.msi"
End If
'**********************************************************************
'** Function; Fetch Property value from MSI database                                     **
'**********************************************************************
Function GetProperty(query)
	GetProperty = ""
	On Error Resume Next
	Set view = db.OpenView("SELECT `Value` FROM Property WHERE `Property` = '"&query&"'") : CheckError
	view.Execute : CheckError
	Set record = view.Fetch : CheckError
	GetProperty = record.StringData(1)
	view.Close
	Set view = nothing
	Set record = nothing
End Function
'**********************************************************************
'** Function; Query MSI database                                     **
'**********************************************************************
Function QueryDatabase(arrOpts)
	On Error Resume Next
	Dim query, file, binary : binary = false
	If LCase(TypeName(arrOpts)) = "string" Then
		query = arrOpts
	Else
		If fs.FileExists(arrOpts(0)) Then
			file = arrOpts(0)
			query = arrOpts(1)
		Else
			query = arrOpts(0)
			file = arrOpts(1)
		End If
		binary = true
	End If
	WScript.Echo query
	If binary Then
		Set record = installer.CreateRecord(1)
		record.SetStream 1, file
	End If
	Set view = db.OpenView (query) : CheckError
	If binary Then
		view.Execute record : CheckError
	Else
		view.Execute : CheckError
	End If
	view.close
	Set view = nothing
	If binary Then Set record = nothing
	binary = false
'	db.commit : CheckError
End Function
'**********************************************************************
'** Subroutine; Check errors in most recently executed MSI command   **
'**********************************************************************
Sub CheckError
	Dim message, errRec
	If Err = 0 Then Exit Sub
	message = Err.Source & " " & Hex(Err) & ": " & Err.Description
	If Not installer Is Nothing Then
		Set errRec = installer.LastErrorRecord
		If Not errRec Is Nothing Then message = message & vbNewLine & errRec.FormatText
	End If
	Wscript.Echo "" : Wscript.Echo message : Wscript.Echo ""
	Wscript.Quit 2
End Sub
'**********************************************************************
'** Function; Push changes to MSI                                    **
'**********************************************************************
Function ProcessMSI(file)
	Set db = installer.OpenDatabase(file, 1)
	On Error Resume Next
	sProperty = GetProperty("ProductVersion")
	If Not Left(sProperty, 2) = "8." Then Exit Function
	Wscript.Echo ""
	sProperty = GetProperty("ProductCPU")
	If LCase(sProperty) = "x86" Then
		QueryDatabase("INSERT INTO `Registry` (`Registry`,`Root`,`Key`,`Name`,`Value`,`Component_`) VALUES ('Servicing_Key_Product_RegKey_7','2','SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\[ProductCode]','DisplayIcon','[SystemFolder]msiexec.exe,0','Servicing_Key_Product')")
	Else
		QueryDatabase("INSERT INTO `Registry` (`Registry`,`Root`,`Key`,`Name`,`Value`,`Component_`) VALUES ('Servicing_Key_Product_RegKey_7','2','SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\[ProductCode]','DisplayIcon','[System64Folder]msiexec.exe,0','Servicing_Key_Product')")
	End If
	sProperty = GetProperty("ProductCode")
	If LCase(sProperty) = "{cbf90bef-21fb-400b-935a-5900785071dd}" Then
		QueryDatabase("UPDATE `Property` SET Value = '{710f4c1c-cc18-4c49-8cbf-51240c89a1a2}' WHERE `Property` = 'ProductCode'")
	End If
	If LCase(sProperty) = "{3aca4f87-8f71-4d1a-bcbe-8c07d3967784}" Then
		QueryDatabase("UPDATE `Property` SET Value = '{ad8a2fa1-06e7-4b0d-927d-6e54b3d31028}' WHERE `Property` = 'ProductCode'")
	End If
	db.commit : CheckError
	Set db = nothing
End Function