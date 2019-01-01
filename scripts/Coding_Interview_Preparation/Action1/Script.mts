

str = read_data_from_text_file_line_by_line
If instr(1, lcase(str), "subfolder") > -1 Then
	msgbox "Yes"
End If

'1. Write a program to read data from a text file
Function read_data_from_text_file_line_by_line()
	
	projectDir = ProjectDirectory()
	testDataPath =  projectDir & "\test_data\testFolder1\testFolder3\testFolder4\testFolder5\testFile1.txt" 
	
	Set fso = CreateObject("Scripting.FileSystemObject")
	set file = fso.OpenTextFile(testDataPath)
	
	Do until file.AtEndOfStream
		line = file.ReadLine
		If line = "" Then
			
		End If
	Loop
	
	Set file = Nothing
	Set fso = Nothing
	
	'read_data_from_text_file_line_by_line = f
	
End Function

exittest

systemutil.CloseProcessByName "chrome.exe"
systemutil.Run "chrome.exe","https://www.facebook.com/", , , 3

wait 5

Set homePage = HomePageInstance()
homePage.setUserName("Sharif")
homePage.setPassword("Password")
homePage.clickLogin()



Set homePage = Nothing

exittest


Const XMLDataFile = "C:\QA\GIT\QTPWebTesting\libraries\ModelRepository\locators.xml"
Set xmlDoc = CreateObject("Microsoft.XMLDOM")
xmlDoc.Async = False
xmlDoc.Load(XMLDataFile)

' Getting the number of Nodes (books)
Set nodes = xmlDoc.SelectNodes("/Pages/Page")
Print "Total Page " & nodes.Length    ' Displays 2

' get all titles
Set nodes = xmlDoc.SelectNodes("/Pages/Page/Name/text()")

' get their values
For i = 0 To (nodes.Length - 1)
   Title = nodes(i).NodeValue
   Print "Title is" & (i + 1) & ": " & Title
Next

Set nodes = Nothing

'lib = Environment("ProductDir")
'msgbox lib
'+ "\bin\Newtonsoft.Json.dll"

'1. Write a program to read data from a text file
'Call read_data_from_text_file_line_by_line()

'2 Write a program to write data into a text file
'Call write_data_to_text_file()

'3 Write a program to print all lines that contains a word either “Print” or “whether”
'call search_specific_text()

'4 Write a program to print the current foldername
'Call getCurrentFolderName()

'5 Write a program to print files in a given folder
'Call print_sub_fileNames_and_sub_folderNames_of_a_given_folder()

'7 Write a program to print all drives in the file system
'Call print_all_drives_names

'8 Write a program to print current drive name
'Call print_current_drive_name()

''9 Print the oldest modified and newest modified date of a file
'Call print_last_modified_file_name_and_date()

'Delete a file with 0 KB size (an empty file)
'Call delete_an_empty_file()


'Call fnSelectCase("")


