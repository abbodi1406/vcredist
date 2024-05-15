' Usage:
'
' cscript <name_of_file>.vbs
'
' script by dumpydooby
' modded by ricktendo64
' updated by abbodi1406
Option Explicit
Dim ws, installer, fs, db, view, record, x, sProperty, icon86, icon64
Set ws = WScript.CreateObject("WScript.Shell")
Set fs = CreateObject("Scripting.FileSystemObject")
Set installer = WScript.CreateObject("WindowsInstaller.Installer")
If WScript.Arguments.Count <> 0 Then
	For each x in WScript.Arguments
		ProcessMSI x
	Next
Else
	If fs.FileExists("vc_runtimeAdditional_x64.msi") Then ProcessMSI "vc_runtimeAdditional_x64.msi"
	If fs.FileExists("vc_runtimeAdditional_x86.msi") Then ProcessMSI "vc_runtimeAdditional_x86.msi"
	If fs.FileExists("vc_runtimeMinimum_x64.msi") Then ProcessMSI "vc_runtimeMinimum_x64.msi"
	If fs.FileExists("vc_runtimeMinimum_x86.msi") Then ProcessMSI "vc_runtimeMinimum_x86.msi"
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
	Wscript.Echo ""
	QueryDatabase("DELETE FROM `CustomAction` WHERE `Action` = 'CA_LaunchCondition_4.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `InstallExecuteSequence` WHERE `Action` = 'CA_LaunchCondition_4.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `InstallUISequence` WHERE `Action` = 'CA_LaunchCondition_4.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `Property` WHERE `Property` = 'ARPSYSTEMCOMPONENT'") 
	QueryDatabase("DELETE FROM `Property` WHERE `Property` = 'ARPNOMODIFY'") 
	QueryDatabase("DELETE FROM `Property` WHERE `Property` = 'ARPNOREPAIR'") 
	QueryDatabase("INSERT INTO `Property` (`Property`,`Value`) VALUES ('ARPNOMODIFY','1')") 
	QueryDatabase("INSERT INTO `Property` (`Property`,`Value`) VALUES ('ARPNOREPAIR','1')") 
	sProperty = GetProperty("ProductVersion")
	If Left(sProperty, 2) = "11" Then
		icon86 = "_x86_VC"
		icon64 = "_amd64_VC"
	Else
		icon86 = ""
		icon64 = ""
	End If
	sProperty = GetProperty("ProductCPU")
	If LCase(sProperty) = "x86" Then
		QueryDatabase("INSERT INTO `Registry` (`Registry`,`Root`,`Key`,`Name`,`Value`,`Component_`) VALUES ('Servicing_Key_ProductEdition_RegKey_9','2','SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\[ProductCode]','DisplayIcon','[SystemFolder"&icon86&"]msiexec.exe,0','Servicing_Key_ProductEdition_x86')") 
	Else
		QueryDatabase("INSERT INTO `Registry` (`Registry`,`Root`,`Key`,`Name`,`Value`,`Component_`) VALUES ('Servicing_Key_ProductEdition_RegKey_9','2','SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\[ProductCode]','DisplayIcon','[System64Folder"&icon64&"]msiexec.exe,0','Servicing_Key_ProductEdition_amd64')") 
	End If
	db.commit : CheckError
	Set db = nothing
End Function