For Each MSIPath in WScript.Arguments
	Set MSIDetails = EvaluateMSI(MSIPath)
	WScript.Echo "   Product Name: " & MSIDetails("ProductName")
	WScript.Echo "   Product Code: " & MSIDetails("ProductCode")
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
	Set EvaluateMSI = objDictionary
	On Error Goto 0
End Function
