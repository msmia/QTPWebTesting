Call Kill_Process("excel.exe")
Call Kill_Process("iexplore.exe")
Call Kill_Process("uft.exe")


Public Function Kill_Process(strProgramName)

	'ex: notepad.exe
	'ex: AcroRd32.exe
	'ex: excel.exe

	Set WMI = GetObject("winmgmts:\\")
	Set allItem = WMI.ExecQuery("Select * from Win32_Process Where Name = "&"'"&strProgramName&"'")	
	
	For Each item in allItem
	   	item.Terminate()
	Next

	

End Function