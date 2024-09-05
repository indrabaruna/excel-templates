Sub unlockPenguji()

Worksheets("DATA PENGUJI").Unprotect ThisWorkbook.Names("password").Value

End Sub

Sub unlockPemicu()

Worksheets("DATA PEMICU").Unprotect ThisWorkbook.Names("password").Value

End Sub

Sub lockPemicu()

' =================Pemicu======================
Dim dataPemicuLastRow As Long
Dim dataPemicuSelectedRange As Range
Dim dataPemicuMainWS As Worksheet
Set dataPemicuMainWS = Worksheets("DATA PEMICU")


dataPemicuLastRow = dataPemicuMainWS.Cells.Find("*", SearchOrder:=xlByRows, _
SearchDirection:=xlPrevious).Row

Set dataPemicuSelectedRange = dataPemicuMainWS.Range("G14:Z" & dataPemicuLastRow)

' ================= Lock Penguji ======================
 
Sheets("DATA PEMICU").Select
Range("G5:X" & dataPemicuLastRow).Select
Selection.Locked = True
 
Sheets("DATA PEMICU").Select
Range("U14:V" & dataPemicuLastRow).Select
Selection.Locked = False

Sheets("DATA PEMICU").Select
Range("G6:M10").Select
Selection.Locked = False
 
dataPemicuMainWS.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowSorting:=True, AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFiltering:=True, Password:=ThisWorkbook.Names("password").Value
  
End Sub
Sub lockPenguji()

' =================Penguji======================
Dim dataPengujiLastRow As Long
Dim dataPengujiSelectedRange As Range
Dim pengujiMainWS As Worksheet

Set dataPengujiMainWS = Worksheets("DATA PENGUJI")

' =================Replace Value======================
' ================= Last Row ======================
dataPengujiLastRow = dataPengujiMainWS.Cells.Find("*", SearchOrder:=xlByRows, _
SearchDirection:=xlPrevious).Row

Set dataPengujiSelectedRange = dataPengujiMainWS.Range("G14:Z" & dataPengujiLastRow)

' ================= Lock Penguji ======================

Sheets("DATA PENGUJI").Select
Range("G5:X" & dataPengujiLastRow).Select
Selection.Locked = True
 
Sheets("DATA PENGUJI").Select
Range("I16:I" & dataPengujiLastRow).Select
Selection.Locked = False

Sheets("DATA PENGUJI").Select
Range("U14:V" & dataPengujiLastRow).Select
Selection.Locked = False

Sheets("DATA PENGUJI").Select
Range("G6:M10").Select
Selection.Locked = False
 
dataPengujiMainWS.Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowSorting:=True, AllowFiltering:=True, Password:=ThisWorkbook.Names("password").Value

End Sub
Sub unlockDetailPenguji()

Worksheets("DPENGUJI DETAIL").Unprotect ThisWorkbook.Names("password").Value

End Sub

Sub unlockDetailPemicu()

Worksheets("DPEMICU DETAIL").Unprotect ThisWorkbook.Names("password").Value

End Sub
Sub lockDetailPenguji()

Worksheets("DPENGUJI DETAIL").Protect Password:=ThisWorkbook.Names("password").Value

End Sub
Sub lockDetailPemicu()

Worksheets("DPEMICU DETAIL").Protect Password:=ThisWorkbook.Names("password").Value

End Sub
Sub lockUploadMainMaster()

Worksheets("UPLOAD_1ST").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=ThisWorkbook.Names("password").Value
Worksheets("UPLOAD_1ST").EnableSelection = xlNoSelection
Worksheets("UPLOAD_2ND").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=ThisWorkbook.Names("password").Value
Worksheets("UPLOAD_2ND").EnableSelection = xlNoSelection
Worksheets("Main Level Data Source").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=ThisWorkbook.Names("password").Value
Worksheets("Main Level Data Source").EnableSelection = xlNoSelection
Worksheets("MASTER").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, Password:=ThisWorkbook.Names("password").Value
Worksheets("MASTER").EnableSelection = xlNoSelection

' ================= Lock Main ======================
  
  Sheets("MAIN LEVEL").Select
  Range("G13:AS103").Select
    Selection.Locked = True
    
  Sheets("MAIN LEVEL").Select
  Range("AE13:AI14").Select
    Selection.Locked = False
 
  Sheets("MAIN LEVEL").Select
  Range("AE54:AI57").Select
    Selection.Locked = False
 
  Sheets("MAIN LEVEL").Select
  Range("AE77:AN80").Select
    Selection.Locked = False
 
  Sheets("MAIN LEVEL").Select
  Range("AE85:AI103").Select
    Selection.Locked = False
 
  Sheets("MAIN LEVEL").Select
  Range("G108:AS120").Select
    Selection.Locked = False
  
   
  Worksheets("MAIN LEVEL").Protect DrawingObjects:=True, Contents:=True, Scenarios:=True, AllowSorting:=False, AllowFormattingCells:=True, AllowFormattingColumns:=True, AllowFiltering:=False, Password:=ThisWorkbook.Names("password").Value


End Sub
Sub unlockUploadMainMaster()

Worksheets("UPLOAD_1ST").Unprotect ThisWorkbook.Names("password").Value
Worksheets("UPLOAD_2ND").Unprotect ThisWorkbook.Names("password").Value
Worksheets("Main Level Data Source").Unprotect ThisWorkbook.Names("password").Value
Worksheets("MASTER").Unprotect ThisWorkbook.Names("password").Value
Worksheets("MAIN LEVEL").Unprotect ThisWorkbook.Names("password").Value

End Sub
Sub auto_close()

pemicupenguji.Name = Sheets("DATA PEMICU").Range("M11").Value & "_" & Sheets("DATA PENGUJI").Range("M10").Value

End Sub
Sub unlockall()

Call LockEWS.unlockDetailPemicu
Call LockEWS.unlockDetailPenguji
Call LockEWS.unlockPemicu
Call LockEWS.unlockPenguji
Call LockEWS.unlockUploadMainMaster

End Sub