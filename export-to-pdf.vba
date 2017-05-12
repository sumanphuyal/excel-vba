Private Sub ExportToPDF_Click()
    Dim wb As Workbook
    Set wb = ThisWorkbook
    'change file path to path where file is located
    Dim MyPath As String
    MyPath = Application.ActiveWorkbook.Path
    ' Change drive/directory to MyPath.
    ChDrive MyPath
    ChDir MyPath
    
    Dim workSheetName As String
    workSheetName = "Sheet1" 'replace this with the worksheet that you want to export
    
    On Error GoTo errHandler
      
    'setup page to export
    With wb.Worksheets(workSheetName).PageSetup
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
    End With

   wb.Worksheets(workSheetName).Activate
   Dim defaultFileName As String
   'I am using worksheet name as default file name, if you want different default file name, please do so in this line.
   defaultFileName = workSheetName 
    
   'select range that you want to export
    wb.Worksheets(workSheetName).UsedRange.Select

    Dim keepAsking As Boolean
    keepAsking = True
    Do While (keepAsking)
        'open file saveas window
        myFile = Application.GetSaveAsFilename _
        (InitialFileName:=defaultFileName, _
            FileFilter:="PDF Files (*.pdf), *.pdf", _
            Title:="Select Folder and FileName to save")
    
     'check if file already exist and if it does let user decide whether to override or exit
        Dim r As String
        Dim config As Long
        config = vbYesNo + vbQuestion
        
        If myFile <> "False" Then
            If Dir(myFile) <> "" Then
                r = MsgBox("A file named """ & myFile & """ already exist in this location. Do you want to replace it?", config)
                If r = vbYes Then
                    Selection.ExportAsFixedFormat _
                    Type:=xlTypePDF, _
                    Filename:=myFile, _
                    Quality:=xlQualityStandard, _
                    IncludeDocProperties:=True, _
                    IgnorePrintAreas:=False, _
                    OpenAfterPublish:=True
                    
                    keepAsking = False
                End If
            Else
                Selection.ExportAsFixedFormat _
                    Type:=xlTypePDF, _
                    Filename:=myFile, _
                    Quality:=xlQualityStandard, _
                    IncludeDocProperties:=True, _
                    IgnorePrintAreas:=False, _
                    OpenAfterPublish:=True
                    
                    keepAsking = False
            End If
        Else 'if Cancel is clicked, exit.
            wb.Worksheets(workSheetName).Select
            
            Exit Sub
        End If
    Loop
    
exitHandler:
    wb.Worksheets(workSheetName).Select
    Exit Sub
errHandler:
    Debug.Print ("Error")
    r = MsgBox("Could not create PDF file. Please check for open files and try again.", vbInformation)
    Resume exitHandler
    
End Sub
