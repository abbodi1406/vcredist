' Regression check for the 7z SFX build scripts.
'
' Each 7zSfx_*.cmd packs an AIO installer and then clears the Archive attribute
' on the EXE it just produced, so the finished artefact is not left marked as
' "needs archiving". The single-file "attrib -A .\<name>" line must therefore
' name the same file the preceding "copy /b ..." line created.
'
' This verifies that invariant for every build script in the folder. It needs
' neither 7-Zip nor any MSI payload, so it runs anywhere cscript.exe does.
'
' usage: cscript //nologo test_attrib_target.vbs [folder]
' exit : 0 = all scripts consistent, 1 = mismatch found, 2 = cannot check

Option Explicit

Dim fso, folder, files, f, txt, lines, i, line
Dim built, target, failures, checked, reCopy, reAttr, hits

Set fso = CreateObject("Scripting.FileSystemObject")

If WScript.Arguments.Count > 0 Then
	folder = WScript.Arguments(0)
Else
	folder = fso.GetParentFolderName(WScript.ScriptFullName)
End If

If Not fso.FolderExists(folder) Then
	WScript.Echo "folder not found: " & folder
	WScript.Quit 2
End If

Set reCopy = New RegExp
reCopy.Pattern = "VisualCppRedist\S*?\.exe"
reCopy.IgnoreCase = False

Set reAttr = New RegExp
reAttr.Pattern = "attrib\s+-A\s+\.\\(VisualCppRedist\S*?\.exe)"
reAttr.IgnoreCase = False

failures = 0
checked = 0

For Each f In fso.GetFolder(folder).Files
	If LCase(Left(f.Name, 6)) = "7zsfx_" And LCase(Right(f.Name, 4)) = ".cmd" Then
		checked = checked + 1
		built = ""
		target = ""

		txt = fso.OpenTextFile(f.Path, 1).ReadAll
		lines = Split(Replace(Replace(txt, vbCrLf, vbLf), vbCr, vbLf), vbLf)

		For i = 0 To UBound(lines)
			line = lines(i)

			' output name = last VisualCppRedist*.exe on the "copy /b" line
			If InStr(1, line, "copy /b", vbTextCompare) > 0 Then
				Set hits = reCopy.Execute(line)
				If hits.Count > 0 Then
					built = hits(hits.Count - 1).Value
				End If
			End If

			' single-file attrib target; skip the bulk "attrib -A .\* /S /D" sweep
			If InStr(1, line, "attrib -A .\", vbTextCompare) > 0 And InStr(1, line, "/S /D", vbTextCompare) = 0 Then
				Set hits = reAttr.Execute(line)
				If hits.Count > 0 Then
					target = hits(0).SubMatches(0)
				End If
			End If
		Next

		If built = "" Then
			WScript.Echo "FAIL " & f.Name & ": no output name on a 'copy /b' line"
			failures = failures + 1
		ElseIf target = "" Then
			WScript.Echo "FAIL " & f.Name & ": no single-file 'attrib -A .\...' line"
			failures = failures + 1
		ElseIf built <> target Then
			WScript.Echo "FAIL " & f.Name
			WScript.Echo "       builds  : " & built
			WScript.Echo "       attribs : " & target & "   <-- never created, Archive bit stays set"
			failures = failures + 1
		Else
			WScript.Echo "ok   " & f.Name & "  (" & built & ")"
		End If
	End If
Next

WScript.Echo ""
If checked = 0 Then
	WScript.Echo "no 7zSfx_*.cmd found in " & folder
	WScript.Quit 2
End If
If failures > 0 Then
	WScript.Echo "RESULT: " & failures & " of " & checked & " build script(s) clear the Archive bit on a file they never built"
	WScript.Quit 1
End If

WScript.Echo "RESULT: all " & checked & " build script(s) agree with their own output name"
WScript.Quit 0
