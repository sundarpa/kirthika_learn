Attribute VB_Name = "Module2"
Sub FetchDataFromNotepad1()
    Dim filePath As String
  Dim searchData As String
    Dim fileContent As String
    Dim result As String
    
    ' Get the file path from the code itself
    filePath = "C:\Users\Lenovo\Desktop\Kirthika\new 1.txt"
    
    ' Get the search term from cell D4
    searchData = Range("D4").Value
    
    ' Check if the file path is provided
    If filePath = "" Then
        MsgBox "File path is not specified."
        Exit Sub
    End If
    
    ' Check if the file exists
    If Dir(filePath) = "" Then
        MsgBox "File not found."
        Exit Sub
    End If
    
    ' Read the content of the Notepad file
    Open filePath For Input As #1
    fileContent = Input(LOF(1), #1)
    Close #1
    
    ' Clear the result string
    result = ""
    
    ' Split the file content by line
    Dim lines As Variant
    lines = Split(fileContent, vbCrLf)
    
    ' Loop through each line in the file
    Dim i As Long
    For i = LBound(lines) To UBound(lines)
        ' Check if the search term is found in the line (main keyword)
        If InStr(1, LCase(lines(i)), LCase(searchData)) > 0 Then
            ' Add the line to the result string
            result = result & lines(i) & vbCrLf
        End If
    Next i
    
    ' Add the data from cell D6 to the result
    result = result & Range("D6").Value
    
    ' Display the result in cell D9
    Range("D9").Value = result
    
    ' Inform the user if no matching inputs found
    If result = "" Then
        MsgBox "No matching inputs found."
    End If
End Sub

