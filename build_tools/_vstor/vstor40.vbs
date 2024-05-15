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
	If fs.FileExists("vstor40_x64.msi") Then ProcessMSI "vstor40_x64.msi"
	If fs.FileExists("vstor40_x86.msi") Then ProcessMSI "vstor40_x86.msi"
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
	If Not Left(sProperty, 3) = "10." Then Exit Function
	Wscript.Echo ""
	QueryDatabase("DELETE FROM `CustomAction` WHERE `Source` = 'CSetupMM_URT_amd64.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `CustomAction` WHERE `Source` = 'CSetupMM_URT_x86.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `Component` WHERE `Directory_` = 'CSetupMM_URT_amd64.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `Component` WHERE `Directory_` = 'CSetupMM_URT_x86.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `CreateFolder` WHERE `Directory_` = 'CSetupMM_URT_amd64.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `CreateFolder` WHERE `Directory_` = 'CSetupMM_URT_x86.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `Directory` WHERE `Directory` = 'CSetupMM_URT_amd64.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `Directory` WHERE `Directory` = 'CSetupMM_URT_x86.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `CreateFolder` WHERE `Component_` = 'GUIH_ARP_TRIR_30871.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `Component` WHERE `Component` = 'GUIH_ARP_TRIR_30871.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `Registry` WHERE `Component_` = 'GUIH_ARP_TRIR_30871.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `FeatureComponents` WHERE `Component_` = 'GUIH_ARP_TRIR_30871.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `FeatureComponents` WHERE `Feature_` = 'TRIN_TRIR_SETUP'") 
	QueryDatabase("DELETE FROM `FeatureExtensionData` WHERE `FeatureId` = 'TRIN_TRIR_SETUP'") 
	QueryDatabase("DELETE FROM `Feature` WHERE `Feature` = 'TRIN_TRIR_SETUP'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'CAB_Setup_for__TRIN_TRIR_ENU_X86_IXP_15354_amd64_ln'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'CAB_Setup_for__TRIN_TRIR_ENU_X86_IXP_15354_x86_ln'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'VS_Setup_MSI__For__TRIN_TRIR_ENU_X86_IXP_12374_amd64_enu'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'VS_Setup_MSI__For__TRIN_TRIR_ENU_X86_IXP_12374_x86_enu'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_INSTALL_EXE_12960_amd64_ln'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_INSTALL_EXE_12960_x86_ln'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_INI_COMP_13899_cn_ln'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_amd64_ara'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_amd64_chs'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_amd64_cht'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_amd64_dan'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_amd64_deu'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_amd64_enu'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_amd64_esn'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_amd64_fin'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_amd64_fra'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_amd64_heb'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_amd64_ita'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_amd64_jpn'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_amd64_kor'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_amd64_nld'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_amd64_nor'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_amd64_plk'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_amd64_ptb'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_amd64_rus'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_amd64_sve'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_x86_ara'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_x86_chs'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_x86_cht'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_x86_dan'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_x86_deu'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_x86_enu'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_x86_esn'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_x86_fin'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_x86_fra'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_x86_heb'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_x86_ita'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_x86_jpn'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_x86_kor'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_x86_nld'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_x86_nor'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_x86_plk'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_x86_ptb'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_x86_rus'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'SETUPUI_VSTO_RESDLL_16893_x86_sve'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'VSTO_EULA_13861_ara'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'VSTO_EULA_13861_chs'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'VSTO_EULA_13861_cht'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'VSTO_EULA_13861_dan'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'VSTO_EULA_13861_deu'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'VSTO_EULA_13861_enu'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'VSTO_EULA_13861_esn'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'VSTO_EULA_13861_fin'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'VSTO_EULA_13861_fra'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'VSTO_EULA_13861_heb'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'VSTO_EULA_13861_ita'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'VSTO_EULA_13861_jpn'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'VSTO_EULA_13861_kor'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'VSTO_EULA_13861_nld'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'VSTO_EULA_13861_nor'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'VSTO_EULA_13861_plk'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'VSTO_EULA_13861_ptb'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'VSTO_EULA_13861_rus'") 
	QueryDatabase("DELETE FROM `File` WHERE `Component_` = 'VSTO_EULA_13861_sve'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'FL_eula_1033_txt_95542_95542_cn_ara.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'FL_eula_1033_txt_95542_95542_cn_chs.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'FL_eula_1033_txt_95542_95542_cn_cht.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'FL_eula_1033_txt_95542_95542_cn_dan.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'FL_eula_1033_txt_95542_95542_cn_deu.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'FL_eula_1033_txt_95542_95542_cn_enu.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'FL_eula_1033_txt_95542_95542_cn_esn.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'FL_eula_1033_txt_95542_95542_cn_fin.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'FL_eula_1033_txt_95542_95542_cn_fra.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'FL_eula_1033_txt_95542_95542_cn_heb.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'FL_eula_1033_txt_95542_95542_cn_ita.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'FL_eula_1033_txt_95542_95542_cn_jpn.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'FL_eula_1033_txt_95542_95542_cn_kor.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'FL_eula_1033_txt_95542_95542_cn_nld.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'FL_eula_1033_txt_95542_95542_cn_nor.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'FL_eula_1033_txt_95542_95542_cn_plk.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'FL_eula_1033_txt_95542_95542_cn_ptb.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'FL_eula_1033_txt_95542_95542_cn_rus.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'FL_eula_1033_txt_95542_95542_cn_sve.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'FL_BASELINE_DAT_95661_95661_cn_ln.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `MsiFileHash` WHERE `File_` = 'FL_globdata_ini_136339_136339_cn_ln.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `CustomAction` WHERE `Action` = 'CA_SETARPSYSTEMCOMPONENT_amd64_enu'") 
	QueryDatabase("DELETE FROM `CustomAction` WHERE `Action` = 'CA_SETARPSYSTEMCOMPONENT_x86_enu'") 
	QueryDatabase("DELETE FROM `InstallExecuteSequence` WHERE `Action` = 'CA_SETARPSYSTEMCOMPONENT_amd64_enu'") 
	QueryDatabase("DELETE FROM `InstallExecuteSequence` WHERE `Action` = 'CA_SETARPSYSTEMCOMPONENT_x86_enu'") 
	QueryDatabase("DELETE FROM `InstallExecuteSequence` WHERE `Action` = 'CA_SetTRIR_Express_Dir_amd64_enu.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `InstallExecuteSequence` WHERE `Action` = 'CA_SetTRIR_Express_Dir_x86_enu.3643236F_FC70_11D3_A536_0090278A1BB8'") 
	QueryDatabase("DELETE FROM `Property` WHERE `Property` = 'MAINTMODE'") 
	sProperty = GetProperty("ProductCPU")
	If LCase(sProperty) = "x86" Then
		QueryDatabase("INSERT INTO `Registry` (`Registry`,`Root`,`Key`,`Name`,`Value`,`Component_`) VALUES ('Servicing_Key_ProductEdition_RegKey_8',2,'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\[ProductCode]','DisplayIcon','[SystemFolder]msiexec.exe,0','Servicing_Key_ProductEdition')") 
	Else
		QueryDatabase("INSERT INTO `Component` (`Component`,`ComponentId`,`Directory_`,`Attributes`,`Condition`,`KeyPath`) VALUES ('Servicing_Key_ProductEdition_amd64','{F01D9F80-E4CF-4940-9A85-9D2C1FB6F943}','TARGETDIR',260,'','Servicing_Key_ProductEdition_RegKey_8')") 
		QueryDatabase("INSERT INTO `Directory` (`Directory`,`Directory_Parent`,`DefaultDir`) VALUES ('System64Folder','WindowsFolder_amd64.3643236F_FC70_11D3_A536_0090278A1BB8','System64')") 
		QueryDatabase("INSERT INTO `FeatureComponents` (`Feature_`,`Component_`) VALUES ('Servicing_Key','Servicing_Key_ProductEdition_amd64')") 
		QueryDatabase("INSERT INTO `Registry` (`Registry`,`Root`,`Key`,`Name`,`Value`,`Component_`) VALUES ('Servicing_Key_ProductEdition_RegKey_8',2,'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\[ProductCode]','DisplayIcon','[System64Folder]msiexec.exe,0','Servicing_Key_ProductEdition_amd64')") 
'		QueryDatabase("UPDATE `CustomAction` SET Target = '[SystemFolder]' WHERE `Action` = 'CA_SystemFolder_amd64.3643236F_FC70_11D3_A536_0090278A1BB8'") 
		QueryDatabase("UPDATE `Registry` SET Value = '[System64Folder]notepad.exe %1' WHERE `Registry` = 'reg9B8D45BEAA1AF7CD505F85B2F0254327'") 
		QueryDatabase("DELETE FROM `Directory` WHERE `Directory` = 'SystemFolder'") 
	End If
	db.commit : CheckError
	Set db = nothing
End Function