Dim objFSO, objFolder, strScriptPath, strNewFileName, strFolderName, strSourceFilePath, strDestinationFilePath, strFileContents

' Create FileSystemObject object
Set objFSO = CreateObject("Scripting.FileSystemObject")

' Get the path to the folder where the script is located
strScriptPath = objFSO.GetParentFolderName(WScript.ScriptFullName)

' Get the name of the current folder (without the full path)
Set objFolder = objFSO.GetFolder(strScriptPath)
strFolderName = objFolder.Name

' Check if Sample.xaml file exists
strSourceFilePath = objFSO.BuildPath(strScriptPath, "Sample.xaml")
If Not objFSO.FileExists(strSourceFilePath) Then
    MsgBox "Couldn't find Sample.xaml file in script folder", vbExclamation, "File Not Found"
    Set objFSO = Nothing
    WScript.Quit
End If

' Validate and input new file name
Do
    strNewFileName = InputBox("Type new file name (without .xaml extension):", "New file")
    
    ' Check if the user clicked Cancel or entered an empty name
    If strNewFileName = "" Then
        Set objFSO = Nothing
        WScript.Quit
    End If
    
    ' Check for invalid characters in the file name
    Dim invalidChars, i
    invalidChars = "\/:*?""<>|"
    Dim isValid
    isValid = True
    
    For i = 1 To Len(invalidChars)
        If InStr(strNewFileName, Mid(invalidChars, i, 1)) > 0 Then
            isValid = False
            Exit For
        End If
    Next
    
    If Not isValid Then
        MsgBox "The file name contains invalid characters. Please provide a valid file name.", vbExclamation, "Invalid File Name"
    End If
Loop While Not isValid

' Continue with the rest of the code...

' If the user provided a file name
If strNewFileName <> "" Then
    ' If the name starts with "#", replace "#" with the folder name as a prefix
    If Left(strNewFileName, 1) = "#" Then
        strNewFileName = Replace(strNewFileName, "#", strFolderName & "_", 1, 1, vbTextCompare)
    End If
    
    ' Path to the destination file
    strDestinationFilePath = objFSO.BuildPath(strScriptPath, strNewFileName & ".xaml")
    
    ' Check if the destination file already exists
    If objFSO.FileExists(strDestinationFilePath) Then
        MsgBox "A file with the name " & strNewFileName & ".xaml already exists.", vbExclamation, "File Already Exists"
        Set objFSO = Nothing
        WScript.Quit
    End If
    
    ' Create a copy of the file
    objFSO.CopyFile strSourceFilePath, strDestinationFilePath
    
    ' Open the copied file to replace names
    Set objFile = objFSO.OpenTextFile(strDestinationFilePath, 1)
    strFileContents = objFile.ReadAll
    objFile.Close
    
    ' Replace all occurrences of "Sample" with the provided file name (Case Insensitive)
    strFileContents = Replace(strFileContents, "Sample", strNewFileName, 1, -1, 1)
    
    ' Save the modified data back to the file
    Set objFile = objFSO.OpenTextFile(strDestinationFilePath, 2)
    objFile.Write strFileContents
    objFile.Close

End If

' Release the FileSystemObject object
Set objFSO = Nothing
