projectDirectory = ProjectDirectory()
print projectDirectory


Function ProjectDirectory()
	
	On error Resume Next
	
	testDir = environment.Value("TestDir")
	projectFolderName = environment.Value("projectFolderName")
	
	If not isEmpty(projectFolderName) Then
	   If instr(1, testDir, projectFolderName) > -1 Then
		  demiliter = split(testDir, projectFolderName)
		  ProjectDirectory = demiliter(0)&projectFolderName
	   End If
	Else
		  reporter.ReportEvent micFail, "Error number: "& err.number& " Error description: " & err.description, "" 
	End If
	
	On Error Goto 0
	
End Function





exittest

ar1 = array("a","b","c")
ar2 = array("d","e","f")

a1 = join(ar1,",")
a2 = join(ar2,",")

joined = a1&","&+a2

asplit = split(joined,",")
msgbox asplit(3)

exittest

systemutil.Run "chrome.exe", "http://newtours.demoaut.com/mercuryreservation.php"

Dim myp,myo,obj


Set myp = Browser("micclass:=Browser").page("micclass:=page")
Set myo = Description.Create
	myo("micclass").value = "WebEdit"
	myo("xpath").value = "//input[@name='userName']"
	
set obj = myp.WebEdit(myo)
	
	If obj.Exist(10) Then
		obj.Set "User Name"
	End If
	
	myo("xpath").value = "//input[@name='password']"
	
	If obj.Exist(10) Then
		obj.Set "Password"
	End If	
	
 systemutil.CloseDescendentProcesses

Set myp = nothing
Set myo = nothing
Set obj = nothing

exittest


 systemutil.Run "chrome.exe", "http://newtours.demoaut.com/mercuryreservation.php"
 @@ hightlight id_;_1507828_;_script infofile_;_ZIP::ssf6.xml_;_
 Set p = Browser("micclass:=Browser").page("micclass:=page")
 
 p.WebEdit("name:=userNam").Set "mercury"
 p.WebEdit("name:=password").Set "mercury"
 p.Image("name:=login").Click
   
 wait 5
   
 systemutil.CloseDescendentProcesses

exittest

Set p = Browser("micclass:=Browser").page("micclass:=page")
'webedit
p.WebEdit("").Set ""
'weblist
p.WebList("").Select ""
'image
p.Image("").Click
'Button
p.webbutton("").Click
'radio button
p.WebRadioGroup("").Select ""


exittest

'Call  the function to Add two Numbers Call Addition(num1,num2) 
a = 10
ab = Addition(a)  
print a
print ab

Function Addition(byval a)  
      a = 50
      Addition = a
End function

exittest

systemutil.Run "chrome.exe","www.Youtube.com", , , 3

Browser("micclass:=Browser").page("micclass:=page").WebElement("innerhtml:=Home").WaitProperty "innerhtml", "Home", 60000

If Browser("micclass:=Browser").page("micclass:=page").WebElement("innerhtml:=Home").Exist(2) Then
	Browser("micclass:=Browser").page("micclass:=page").WebElement("innerhtml:=Home").highlight
	msgbox "Yes"
else
	msgbox "No"
End If

systemutil.CloseDescendentProcesses

exittest

'ArrayList with sort
Set MyList = CreateObject("System.Collections.ArrayList")
MyList.Add("ListItem5")
MyList.Add("ListItem3")
MyList.Add("ListItem2")
MyList.Add("ListItem1")
MyList.Add("ListItem4")

MyList.sort()

For each x in MyList
	msgbox x
Next


Set MyList = Nothing

exittest

' Vartype and typeName
str = "initiation a c b d g sf afafag"
msgbox vartype(str)
msgbox typeName(str)

exittest


str = "initiation a c b d g sf afafag"
spl = split(str," ")
For i = 0 to ubound(spl) - 1
	print spl(i)
Next

exittest

str = "initiation"
match = 0
For i = 1 To len(str)
	str_i = mid(str, i, 1)
	If str_i = "i" Then
		match = match + 1
	End If
 
Next

msgbox match

exittest

'FileSystemObject
Set fso = CreateObject("Scripting.FileSystemObject")
if fso.FileExists("C:\Users\Tuhin\Documents\sql installation1.txt") Then
	Set content = fso.OpenTextFile("C:\Users\Tuhin\Documents\sql installation1.txt")
	
	Do  until content.AtEndOfStream 
		sContent = content.ReadLine
		print sContent
	Loop
End  If

Set content = nothing
Set fso = nothing

exittest


print "Call all the actions in a different test to practice reusability"

RunAction "Driver [ReusableActions]", oneIteration

RunAction "sginin [ReusableActions]", oneIteration

RunAction "placeOrder [ReusableActions]", oneIteration

RunAction "editOrder", oneIteration

RunAction "signout [ReusableActions]", oneIteration



