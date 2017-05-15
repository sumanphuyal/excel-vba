Private Sub ExportToExcel_Click()
  loopAgain:
     'change file path to path where file is located
     Dim MyPath As String
     MyPath = Application.ActiveWorkbook.Path
     ' Change drive/directory to MyPath.
     ChDrive MyPath
     ChDir MyPath
     
    Dim wb As Workbook
    Dim wbNew As Workbook
    Dim sh As Worksheet
    Dim shNew As Worksheet
    Dim r As String

    Dim sheetArray As Variant
    'Update this array with worksheets you want to export.
    'They will be exported in the order as they appear in the array.
    sheetArray = Array("SheetA", "SheetB", "SheetC")
    Set wb = ThisWorkbook
   
    On Error Resume Next
    
    Dim defaultFileName As String
    defaultFileName = "exportFileName" ' update this with your default filename
    
    Dim keepAsking As Boolean
    keepAsking = True
    Do While (keepAsking)
        Dim myFile As String
        myFile = Application.GetSaveAsFilename _
            (InitialFileName:=defaultFileName, _
            FileFilter:="Excel Files (*.xlsx), *.xlsx", _
            Title:="Select Folder and FileName to save")
        'check if file already exist and if it does let user decide whether to override or exit
        Dim config As Long
        config = vbYesNo + vbQuestion
    
        If myFile <> "False" Then
            If Dir(myFile) <> "" Then
                r = MsgBox("A file named """ & myFile & """ already exist in this location. Do you want to replace it?", config)
                If r = vbNo Then
                    GoTo loopAgain
                End If
            End If
            Workbooks.Add ' Open a new workbook
            Set wbNew = ActiveWorkbook

            For Each wSheet In sheetArray
             Set sh = ThisWorkbook.Worksheets(wSheet)
               sh.UsedRange.Copy
         
               'add new sheet into new workbook with the same name
               With wbNew.Worksheets
                   Set shNew = Nothing
                   Set shNew = .Item(sh.Name)
                   If shNew Is Nothing Then
                       .Add After:=.Item(.count)
                       .Item(.count).Name = sh.Name
                       Set shNew = .Item(.count)
                   End If
               End With
         
                 'paste to new workbook/worksheet
                 Application.DisplayAlerts = False
                 With shNew.Range("A1")
                     .PasteSpecial (xlValues)
                     .PasteSpecial (xlFormats)
                     .PasteSpecial (xlPasteColumnWidths)
                 End With
                 
                 'copy all the pictures from source to destination
                 Dim pic As Shape, rng As Range
                 For Each pic In sh.Shapes
                    If pic.Type = msoPicture Then
                    pic.Copy
                        With shNew
                            .Select
                            .Range(pic.TopLeftCell.Address).Select
                            .Paste
                        End With
                        Selection.Placement = xlMoveAndSize
                    End If
                 Next pic
                 Application.DisplayAlerts = True
            Next
            
            'delete the first sheet i.e Sheet 1 that excel insert by default
            Application.DisplayAlerts = False
            wbNew.Worksheets("Sheet1").Delete
            Application.DisplayAlerts = True
            wbNew.SaveAs fileName:=myFile
            wbNew.Close
            
            r = MsgBox(myFile & " successfully exported.", vbInformation)
            keepAsking = False
        Else 'if cancel is pressed exit Sub
            wb.Worksheets(worksheetName).Select
            Exit Sub
        End If
    Loop
    
  exitHandler:
    Exit Sub
  errHandler:
    r = MsgBox("Could not export to excel.", vbInformation)
    Resume exitHandler
End Sub
