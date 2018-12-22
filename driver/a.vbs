
pageObjects = "C:\QA\GIT\QTPWebTesting\page_objects" 

set a = listOfFiles(pageObjects)
for i = 0 to a.count - 1
	msgbox a(i)
next


Function listOfFiles(strDirectory)
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	set libFolder = fso.GetFolder(strDirectory)
	
	set filesArrayList = CreateObject("System.Collections.ArrayList")
	
	For each file in libFolder.files
		filesArrayList.add(file.name)
	Next

	Set listOfFiles = filesArrayList
	
	Set filesArrayList = Nothing
	Set libFolder = Nothing
	Set fso = Nothing
	
End Function




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