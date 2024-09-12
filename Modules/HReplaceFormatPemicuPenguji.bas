Sub vbaReplaceFormatPemicuPenguji(Optional x As Integer)

' =================Penguji======================
Dim pengujiLastRow As Long
Dim pengujiSelectedRange As Range
Dim pengujiMainWS As Worksheet

Set pengujiMainWS = Worksheets("DATA PENGUJI")

' =================Replace Value======================
pengujiMainWS.Columns("T").Replace _
 What:="{", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True

pengujiMainWS.Columns("T").Replace _
 What:="}", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True

pengujiMainWS.Columns("T").Replace _
 What:=",", Replacement:="," & Chr(10), _
 SearchOrder:=xlByColumns, MatchCase:=True

  pengujiMainWS.Columns("T").Replace _
 What:=Chr(34), Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
 
 pengujiMainWS.Columns("T").Replace _
 What:="NPWP", Replacement:=" NPWP", _
 SearchOrder:=xlByColumns, MatchCase:=True

' ================= Last Row ======================
pengujiLastRow = pengujiMainWS.Cells.Find("*", SearchOrder:=xlByRows, _
SearchDirection:=xlPrevious).Row

Set pengujiSelectedRange = pengujiMainWS.Range("G14:Z" & pengujiLastRow)

' ================= Set Formatting ======================

pengujiSelectedRange.HorizontalAlignment = xlLeft
pengujiSelectedRange.VerticalAlignment = xlTop
pengujiSelectedRange.WrapText = True
pengujiMainWS.Range("G:K").EntireColumn.AutoFit
pengujiMainWS.Range("S:W").EntireColumn.AutoFit
    Sheets("DATA PENGUJI").Select
        Range("U16:U" & pengujiLastRow).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:= _
        "Digunakan, Tidak Digunakan - Data Tidak Sesuai,Tidak Digunakan - Beririsan,Tidak Digunakan - Data Sudah Digunakan Sebelumnya"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    Range("W16:W" & pengujiLastRow).FormulaR1C1 = "=IF(RC[-2]=""Digunakan"",RC[-4],RC[-2])"

        Sheets("DATA PENGUJI").Select
        Range("I16:I" & pengujiLastRow).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=MASTER!$G$2:$G$48"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    

'' =================Pemicu======================
'' =================Pemicu======================
Dim pemicuLastRow As Long
Dim pemicuSelectedRange As Range
Dim PemicuMainWS As Worksheet

Set PemicuMainWS = Worksheets("DATA PEMICU")

PemicuMainWS.Columns("T").Replace _
 What:="{", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True

PemicuMainWS.Columns("T").Replace _
 What:="}", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
 

PemicuMainWS.Columns("T").Replace _
 What:=",", Replacement:="," & Chr(10), _
 SearchOrder:=xlByColumns, MatchCase:=True

PemicuMainWS.Columns("T").Replace _
 What:=Chr(34), Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
 
PemicuMainWS.Columns("T").Replace _
 What:="NPWP", Replacement:=" NPWP", _
 SearchOrder:=xlByColumns, MatchCase:=True
 

' ' ================= Set Last Row ======================
'
pemicuLastRow = PemicuMainWS.Cells.Find("*", SearchOrder:=xlByRows, _
SearchDirection:=xlPrevious).Row

Set pemicuSelectedRange = PemicuMainWS.Range("G14:Z" & pemicuLastRow)

 ' ================= Set Formatting ======================
pemicuSelectedRange.HorizontalAlignment = xlLeft
pemicuSelectedRange.VerticalAlignment = xlTop
pemicuSelectedRange.WrapText = True
PemicuMainWS.Range("G:K").EntireColumn.AutoFit
PemicuMainWS.Range("S:W").EntireColumn.AutoFit

    Sheets("DATA PEMICU").Select
        Range("U16:U" & pengujiLastRow).Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:= _
        "Digunakan, Tidak Digunakan - Data Tidak Sesuai,Tidak Digunakan - Beririsan,Tidak Digunakan - Data Sudah Digunakan Sebelumnya"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    Range("W16:W" & pemicuLastRow).FormulaR1C1 = "=IF(RC[-2]=""Digunakan"",RC[-4],RC[-2])"
' =================Detail Penguji======================
' =================Detail Penguji======================
Dim detailpengujiLastRow As Long
Dim detailpengujiSelectedRange As Range
Dim detailpengujiMainWS As Worksheet

Set detailpengujiMainWS = Worksheets("DPENGUJI DETAIL")

' =================Replace Value======================
detailpengujiMainWS.Columns("J").Replace _
 What:="{", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True

detailpengujiMainWS.Columns("J").Replace _
 What:="}", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True

detailpengujiMainWS.Columns("J").Replace _
 What:=",", Replacement:="," & Chr(10), _
 SearchOrder:=xlByColumns, MatchCase:=True

detailpengujiMainWS.Columns("J").Replace _
 What:=Chr(34), Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
 
  detailpengujiMainWS.Columns("J").Replace _
 What:="NPWP", Replacement:=" NPWP", _
 SearchOrder:=xlByColumns, MatchCase:=True

' ================= Last Row ======================
detailpengujiLastRow = detailpengujiMainWS.Cells.Find("*", SearchOrder:=xlByRows, _
SearchDirection:=xlPrevious).Row

Set detailpengujiSelectedRange = detailpengujiMainWS.Range("G14:Z" & detailpengujiLastRow)

' ================= Set Formatting ======================

detailpengujiSelectedRange.HorizontalAlignment = xlLeft
detailpengujiSelectedRange.VerticalAlignment = xlTop
detailpengujiSelectedRange.WrapText = True
detailpengujiSelectedRange.EntireColumn.AutoFit


' =================Pemicu======================
Dim detailpemicuLastRow As Long
Dim detailpemicuSelectedRange As Range
Dim detailPemicuMainWS As Worksheet

Set detailPemicuMainWS = Worksheets("DPEMICU DETAIL")

detailPemicuMainWS.Columns("J").Replace _
 What:="{", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True

detailPemicuMainWS.Columns("J").Replace _
 What:="}", Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True

detailPemicuMainWS.Columns("J").Replace _
 What:=",", Replacement:="," & Chr(10), _
 SearchOrder:=xlByColumns, MatchCase:=True

detailPemicuMainWS.Columns("J").Replace _
 What:=Chr(34), Replacement:="", _
 SearchOrder:=xlByColumns, MatchCase:=True
 
   detailPemicuMainWS.Columns("J").Replace _
 What:="NPWP", Replacement:=" NPWP", _
 SearchOrder:=xlByColumns, MatchCase:=True

' ================= Set Last Row ======================
detailpemicuLastRow = detailPemicuMainWS.Cells.Find("*", SearchOrder:=xlByRows, _
SearchDirection:=xlPrevious).Row

Set detailpemicuSelectedRange = detailPemicuMainWS.Range("G14:Z" & detailpemicuLastRow)

' ================= Set Formatting ======================
detailpemicuSelectedRange.HorizontalAlignment = xlLeft
detailpemicuSelectedRange.VerticalAlignment = xlTop
detailpemicuSelectedRange.WrapText = True
detailpemicuSelectedRange.EntireColumn.AutoFit


' ================= Set Data Validation Formatting ======================
Call dataValidation.dataValidationPengujian_B1
Call dataValidation.dataValidationPengujian_B2
Call dataValidation.dataValidationPengujian_B3
Call dataValidation.dataValidationPengujian_B4
Call dataValidation.dataValidationPengujian_B5
Call dataValidation.dataValidationPengujian_C1
Call dataValidation.dataValidationPengujian_C2
Call dataValidation.dataValidationPengujian_C3
Call dataValidation.dataValidationPengujian_C4
Call dataValidation.dataValidationPengujian_C5
Call dataValidation.dataValidationPengujian_C6
Call dataValidation.dataValidationPengujian_C7
Call dataValidation.dataValidationPengujian_C8
Call dataValidation.dataValidationPengujian_C9
Call dataValidation.dataValidationPengujian_D
Call dataValidation.dataValidationPengujian_E1
Call dataValidation.dataValidationPengujian_E2
Call dataValidation.dataValidationPengujian_E3
Call dataValidation.dataValidationPengujian_E4
Call dataValidation.dataValidationPengujian_E5
Call dataValidation.dataValidationPengujian_E6
Call dataValidation.dataValidationPengujian_E7
Call dataValidation.dataValidationPengujian_E8
Call dataValidation.dataValidationPengujian_E9
Call dataValidation.dataValidationPengujian_E10
Call dataValidation.dataValidationPengujian_E11
Call dataValidation.dataValidationPengujian_E12
Call dataValidation.dataValidationPengujian_E13
Call dataValidation.dataValidationPengujian_F1
Call dataValidation.dataValidationPengujian_F2
Call dataValidation.dataValidationPengujian_F3
Call dataValidation.dataValidationPengujian_F4
Call dataValidation.dataValidationPengujian_G1
Call dataValidation.dataValidationPengujian_G2
Call dataValidation.dataValidationPengujian_G3
Call dataValidation.dataValidationPengujian_G4
Call dataValidation.dataValidationPengujian_H
Call dataValidation.dataValidationPengujian_I
Call dataValidation.dataValidationPengujian_J
Call dataValidation.dataValidationPengujian_K
Call dataValidation.dataValidationPengujian_L
Call dataValidation.dataValidationPengujian_M1
Call dataValidation.dataValidationPengujian_M2
Call dataValidation.dataValidationPengujian_M3
Call dataValidation.dataValidationPengujian_P
Call dataValidation.dataValidationPengujian_Q


' ================= Lock ======================

Call LockEWS.lockDetailPemicu
Call LockEWS.lockDetailPenguji
Call LockEWS.lockPemicu
Call LockEWS.lockPenguji
Call LockEWS.lockUploadMainMaster

        
End Sub