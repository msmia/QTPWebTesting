

Call ProjectDirectory()


Function ProjectDirectory()
	
	On error Resume Next

	scriptdir = CreateObject("Scripting.FileSystemObject").GetParentFolderName(WScript.ScriptFullName)
	projectFolderName = "QTPWebTesting"
	
	If not isEmpty(projectFolderName) Then
	   If instr(1, scriptdir, projectFolderName) > -1 Then
		  demiliter = split(scriptdir, projectFolderName)
		  ProjectDirectory = demiliter(0) & projectFolderName
	   End If
	Else
		  reporter.ReportEvent micFail, "Error number: "& err.number& " Error description: " & err.description, "" 
	End If
	
	On Error Goto 0
	
End Function