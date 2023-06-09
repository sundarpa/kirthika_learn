Attribute VB_Name = "Module1"
Sub FetchDataFromNotepad()
    Dim filePath As String
    Dim searchData As String
    Dim fileContent As String
    Dim result As String
    
    ' Set the file path
    filePath = "C:\Users\Lenovo\Desktop\Kirthika\new 1.txt"
    
    ' Get the search term from cell D4
    searchData = Range("D4").Value
    
    ' Check if the search term is empty
    If Len(searchData) = 0 Then
        MsgBox "Please enter a search term."
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
        ' Check if the line contains a similar match to the search term
        If LCase(lines(i)) Like "*" & LCase(searchData) & "*" Then
            ' Add the line to the result string
            result = result & lines(i) & vbCrLf
        End If
    Next i
    
    ' Display the result in cell D9
    Range("D9").Value = result
    
    ' Inform the user if no matching inputs found
    If result = "" Then
        MsgBox "No matching inputs found."
    End If
End Sub

Sub AddData()
    ' Get the existing result in cell D9
    Dim existingResult As String
    existingResult = Range("D9").Value
    
    ' Get the data from cell D6
    Dim additionalData As String
    additionalData = Range("D6").Value
    
    ' Append the additional data to the existing result
    Dim updatedResult As String
    updatedResult = existingResult & additionalData
    
    ' Update the result in cell D9
    Range("D9").Value = updatedResult
End Sub

Sub SaveResult()
    ' Get the result from cell D9
    Dim result As String
    result = Range("D9").Value
    
    ' Check if the result is empty
    If result = "" Then
        MsgBox "Nothing to save. Result is empty."
        Exit Sub
    End If
    
    ' Specify the file path to save the result
    Dim filePath As String
    filePath = "C:\Users\Lenovo\Desktop\Kirthika\result.txt"
    
    ' Save the result to the file
    Open filePath For Output As #1
    Print #1, result
    Close #1
    
    ' Inform the user about the successful save
    MsgBox "Result saved successfully to: " & filePath
End Sub


Sub RemoveData()
    On Error GoTo ErrorHandler
    
    Dim filePath As String
    filePath = "C:\Users\Lenovo\Desktop\Kirthika\result.txt" ' Specify the file path here
    
    ' Clear the data in cell D6
    Dim removedData As String
    removedData = Range("D6").Value
    Range("D6").ClearContents
    
    ' Check if the file path is provided
    If filePath = "" Then
        MsgBox "File path not available."
        Exit Sub
    End If
    
    ' Check if the file exists
    If Dir(filePath) = "" Then
        MsgBox "File not found."
        Exit Sub
    End If
    
    ' Read the content of the file
    Dim fileContent As String
    Open filePath For Input As #1
    fileContent = Input$(LOF(1), #1)
    Close #1
    
    ' Remove the data from the file content
    fileContent = Replace(fileContent, removedData, "")
    
    ' Write the updated content back to the file
    Open filePath For Output As #1
    Print #1, fileContent
    Close #1
    
    ' Clear the result in cell D9
    Dim result As String
    result = Range("D9").Value
    result = Replace(result, removedData, "")
    Range("D9").Value = result
    
    ' Display a message
    MsgBox "Data removed from the file and result."
    
    Exit Sub

ErrorHandler:
    MsgBox "An error occurred while processing the file."
End Sub

Sub ClearData()
    If MsgBox("Are you sure you want to clear the data?", vbYesNo + vbQuestion) = vbYes Then
        Range("D4").ClearContents ' Clear the data in cell D4
        Range("D9").Value = "" ' Clear the result in cell D9
        Range("D6").ClearContents ' Clear the data in cell D6
    End If
End Sub

