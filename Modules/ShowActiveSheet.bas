Sub VeryHiddenActiveSheet()
  Sheets("MASTER").Visible = xlSheetVeryHidden
  Sheets("UPLOAD_1ST").Visible = xlSheetVeryHidden
  Sheets("UPLOAD_2ND").Visible = xlSheetVeryHidden
  Sheets("Main Level Data Source").Visible = xlSheetVeryHidden
' Sheets("Internal Data Source").Visible = xlSheetVeryHidden
End Sub

Sub ShowActiveSheet()
  Sheets("MASTER").Visible = True
  Sheets("UPLOAD_1ST").Visible = True
  Sheets("UPLOAD_2ND").Visible = True
  Sheets("Main Level Data Source").Visible = True
' Sheets("Internal Data Source").Visible = True

End Sub