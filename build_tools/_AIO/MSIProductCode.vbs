For Each MSIPath in WScript.Arguments
	Set MSIDetails = EvaluateMSI(MSIPath)
	WScript.Echo "   Product Name: " & MSIDetails("ProductName")
	WScript.Echo "   Product Code: " & MSIDetails("ProductCode")
	WScript.Echo "Compressed  : " & MSIDetails("CompressedCode")
	WScript.Echo ""
Next

Function EvaluateMSI(MSIPath)
	On Error Resume Next
	Set oInstaller = CreateObject("WindowsInstaller.Installer")
	Set oDatabase = oInstaller.OpenDatabase(MSIPath, 0)
	Set objDictionary = CreateObject("Scripting.Dictionary")
	Set View = oDatabase.OpenView("Select `Value` From Property WHERE `Property`='ProductName'")
	View.Execute
	Set ProductName = View.Fetch
	objDictionary("ProductName") = ProductName.StringData(1)
	Set View = oDatabase.OpenView("Select `Value` From Property WHERE `Property`='ProductCode'")
	View.Execute
	Set ProductCode = View.Fetch
	objDictionary("ProductCode") = ProductCode.StringData(1)
	objDictionary("CompressedCode") = GetCompressedGuid(ProductCode.StringData(1))
	Set EvaluateMSI = objDictionary
	On Error Goto 0
End Function

Function GetCompressedGuid(sGuid)
    Dim sCompGUID
    Dim i
    sCompGUID = StrReverse(Mid(sGuid, 2, 8))  & _
                StrReverse(Mid(sGuid, 11, 4)) & _
                StrReverse(Mid(sGuid, 16, 4)) 
    For i = 21 To 24
	    If i Mod 2 Then
		    sCompGUID = sCompGUID & Mid(sGuid, (i + 1), 1)
	    Else
		    sCompGUID = sCompGUID & Mid(sGuid, (i - 1), 1)
	    End If
    Next
    For i = 26 To 37
	    If i Mod 2 Then
		    sCompGUID = sCompGUID & Mid(sGuid, (i - 1), 1)
	    Else
		    sCompGUID = sCompGUID & Mid(sGuid, (i + 1), 1)
	    End If
    Next
    GetCompressedGuid = sCompGUID
End Function
