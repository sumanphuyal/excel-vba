'function to protect all worksheets
Private Sub protectAllWorksheets()  
  Dim ws As Worksheet
  Dim pwd As String

  pwd = "abc" ' If you want to protect worksheets without any password use pwd = ""
  For Each ws In Worksheets
      ws.Protect Password:=pwd
  Next ws
End Sub


'function to unprotect all the worksheets
Private Sub unprotectAllWorksheets()  
   Dim ws As Worksheet
   Dim pwd As String
    
pwd = "abc" ' if worksheets were protected without any password use pwd = "" 
   For Each ws In Worksheets
      ws.Unprotect Password:=pwd
   Next ws
End Sub
