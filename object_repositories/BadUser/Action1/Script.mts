Dim sRootFolder: sRootFolder = Environment.Value("TestDir")
splitRootFolderBySlash = split(sRootFolder,"\")
print splitRootFolderBySlash(4)
Environment.Value(T)

sFrameworkFolder = "H:\quick_test_result_practice_UFT\"
sTestCaseFolder = sFrameworkFolder & "TestCases\"
sQTPResultsPathOrig = sFrameworkFolder & "Results\DetailedQTPResults\"
sBatchRunPath = sFrameworkFolder & "Results\SummarizedResults\"
sBatchSheetPath = sFrameworkFolder & "TestCaseNames.xlsx"
sBatchSheetName = "TC_ID"
sResultSheetName = "Sheet1"



exittest
systemutil.Run "chrome.exe","https://www.yahoo.com/"
wait (15)
systemutil.CloseDescendentProcesses
