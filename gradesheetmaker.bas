Attribute VB_Name = "Gradesheetmaker"
Sub GenerateGradeSheets()
        On Error GoTo ErrorHandler
    
    Dim wbTemplate As Workbook
    Dim wsTemplate As Worksheet
    Dim wbData As Workbook
    Dim wsData As Worksheet
    Dim studentName As String
    Dim studentID As String
    Dim newFileName As String
    Dim destinationPath As String
    
    ' Set the path of the template file
    Dim templatePath As String
    ' to be edited
    templatePath = "C:\path\to\file\template.xlsx"
    
    ' Set the path of the data file containing student names and IDs
    Dim dataFilePath As String
    dataFilePath = "C:\path\to\file\list_of_students.xlsx"
    
    ' Set the destination path where the generated grade sheets will be saved
    destinationPath = "C:\path\to\destination\"
    
   ' Open the data file containing student names and IDs
    Set wbData = Workbooks.Open(dataFilePath)
    Set wsData = wbData.Sheets(1)
    
    ' Loop through the list of students in the data file
    Dim lastRow As Long
    lastRow = wsData.Cells(wsData.Rows.Count, "A").End(xlUp).Row
    
    For i = 2 To lastRow ' Assuming student names start from row 2, change if needed
        ' Open the template file
        Set wbTemplate = Workbooks.Open(templatePath)
        Set wsTemplate = wbTemplate.Sheets(1)
        
        ' Get student name and ID from data file
        studentName = wsData.Cells(i, 1).Value
        studentID = wsData.Cells(i, 2).Value
        
        ' Set the name and ID in the template file
        On Error Resume Next
        wsTemplate.Range("A6").Value = studentName
        wsTemplate.Range("C6").Value = studentID
        On Error GoTo 0
        
        If Err.Number <> 0 Then
            MsgBox "Error setting values in the template file: " & Err.Description
            wbTemplate.Close SaveChanges:=False
            Exit Sub
        End If
        
        ' Generate a unique file name for each student
        newFileName = studentName & "_TtRtM_gradesheet.xlsx"
        
        ' Check if a file with the same name already exists in the destination folder
        If FileExists(destinationPath & newFileName) Then
            MsgBox "A file with the name '" & newFileName & "' already exists in the destination folder."
            wbTemplate.Close SaveChanges:=False
            Exit Sub
        End If
        
        ' Save the gradesheet with the student's name in the desired destination folder
        wbTemplate.SaveCopyAs destinationPath & newFileName
        
        ' Close the newly created gradesheet
        wbTemplate.Close SaveChanges:=False
    Next i
    
    ' Close the data file
    wbData.Close SaveChanges:=False
    
    ' Release object references
    Set wsTemplate = Nothing
    Set wbTemplate = Nothing
    Set wsData = Nothing
    Set wbData = Nothing
    
    MsgBox "Grade sheets generated successfully!"
    Exit Sub
    
ErrorHandler:
    MsgBox "An error occurred: " & Err.Description
End Sub

Function FileExists(filePath As String) As Boolean
    FileExists = (Dir(filePath) <> "")
End Function
