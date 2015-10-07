Dim fso
Dim curDir
Dim WinScriptHost
Set fso = CreateObject("Scripting.FileSystemObject")
curDir = fso.GetAbsolutePathName(".")
Set fso = Nothing
Set xlObj = CreateObject("Excel.application")
xlObj.Workbooks.Open curDir & "\Production Starter.xlsm"
xlObj.Run "open_form"
