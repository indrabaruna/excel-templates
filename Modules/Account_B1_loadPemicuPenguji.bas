Sub Account_B1_loadPemicuAndPenguji()

Call FilteringPemicu.PemicuClearFilter
Call FilteringPenguji.PengujiClearFilter
Call LockEWS.unlockPemicu
Call LockEWS.unlockPenguji

Dim pengujiWS As Worksheet, MainWS As Worksheet, Lv2WS As Worksheet
Dim pengujiCopyRng As Range
Dim pengujiLastRow As Long
Dim pengujiFirstRowDestination As Integer

Dim pengujiCriteriaColumnNo_1 As Integer
Dim pengujiCriteriaColumnNo_2 As Integer
Dim firstTabelToBeFilterRow As Integer

Dim pengujiResultLocation1 As String
Dim pengujiResultLocation2 As String
Dim pengujiResultLocation3 As String
Dim pengujiResultLocation4 As String
Dim pengujiResultLocation5 As String
Dim AccountUsed As String

Dim ResultColumn1 As String
Dim ResultColumn2 As String
Dim ResultColumn3 As String
Dim ResultColumn4 As String
Dim ResultColumn5 As String

pengujiCriteriaColumnNo_1 = 3
pengujiCriteriaColumnNo_2 = 15
firstTabelToBeFilterRow = 15

'======================================= Destination ==================================

pengujiFirstRowDestination = 30
ResultColumn1 = "H"
ResultColumn2 = "O"
ResultColumn3 = "K"
ResultColumn4 = "J"
ResultColumn5 = "X"

AccountUsed = "Penjualan Lokal"

pengujiResultLocation1 = ResultColumn1 & firstTabelToBeFilterRow + 1 & ":" & ResultColumn1
pengujiResultLocation2 = ResultColumn2 & firstTabelToBeFilterRow + 1 & ":" & ResultColumn2
pengujiResultLocation3 = ResultColumn3 & firstTabelToBeFilterRow + 1 & ":" & ResultColumn3
pengujiResultLocation4 = ResultColumn4 & firstTabelToBeFilterRow + 1 & ":" & ResultColumn4
pengujiResultLocation5 = ResultColumn5 & firstTabelToBeFilterRow + 1 & ":" & ResultColumn5

Set pengujiWS = Sheets("DATA PENGUJI")
Set MainWS = Sheets("B-1-1-9")
Set Lv2WS = Sheets("B-1")

pengujiLastRow = pengujiWS.Cells.Find("*", SearchOrder:=xlByRows, _
SearchDirection:=xlPrevious).Row

' ======================================= Copy empty template ==================================
    MainWS.Rows("29:9999").Delete
    Sheets("MASTER").Range("AB10:AI25").Copy MainWS.Range("G29")
    
    MainWS.UsedRange.Replace What:="COUNTA($H$29:$H$30))", Replacement:="COUNTA($H$29:$H$31))", LookAt:= _
    xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    FormulaVersion:=xlReplaceFormula2
    
    Lv2WS.UsedRange.Replace What:="!#REF!", Replacement:="!$L$34", LookAt:= _
    xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
    FormulaVersion:=xlReplaceFormula2

'============================Copy Data Pemicu to other Sheet to remove Formula=============================

pengujiWS.Range("W15:W" & pengujiLastRow).Copy
pengujiWS.Range("X15:X" & pengujiLastRow).PasteSpecial Paste:=xlPasteValues
pengujiWS.Columns("X:X").NumberFormat = _
        "_([$Rp-id-ID]* #,##0_);_([$Rp-id-ID]* (#,##0);_([$Rp-id-ID]* ""-""_);_(@_)"


'============================Filtering and Copy Process=============================
With pengujiWS.Range("G15:X" & pengujiLastRow)
    On Error Resume Next
    .AutoFilter field:=pengujiCriteriaColumnNo_1, Criteria1:=AccountUsed
    .AutoFilter field:=pengujiCriteriaColumnNo_2, Criteria1:="Digunakan"
    If pengujiWS.Range(pengujiResultLocation1 & pengujiLastRow).SpecialCells(xlCellTypeVisible).Cells.Count = 0 Then
            MainWS.Range("H" & pengujiFirstRowDestination) = "Data Penguji Tidak Tersedia"
            
    ElseIf pengujiWS.Range(pengujiResultLocation1 & pengujiLastRow).SpecialCells(xlCellTypeVisible).Cells.Count = 1 Then
            pengujiWS.Range(pengujiResultLocation1 & pengujiLastRow).SpecialCells(xlCellTypeVisible).Copy MainWS.Range("H" & pengujiFirstRowDestination)
            pengujiWS.Range(pengujiResultLocation2 & pengujiLastRow).SpecialCells(xlCellTypeVisible).Copy MainWS.Range("I" & pengujiFirstRowDestination)
            pengujiWS.Range(pengujiResultLocation3 & pengujiLastRow).SpecialCells(xlCellTypeVisible).Copy MainWS.Range("J" & pengujiFirstRowDestination)
            pengujiWS.Range(pengujiResultLocation4 & pengujiLastRow).SpecialCells(xlCellTypeVisible).Copy MainWS.Range("K" & pengujiFirstRowDestination)
            pengujiWS.Range(pengujiResultLocation5 & pengujiLastRow).SpecialCells(xlCellTypeVisible).Copy MainWS.Range("L" & pengujiFirstRowDestination)
            
            
    ElseIf pengujiWS.Range(pengujiResultLocation1 & pengujiLastRow).SpecialCells(xlCellTypeVisible).Cells.Count > 1 Then
            MainWS.Range(pengujiFirstRowDestination + 1 & ":" & pengujiWS.AutoFilter.Range.Columns(firstTabelToBeFilterRow).SpecialCells(xlCellTypeVisible).Cells.Count + pengujiFirstRowDestination - 1).Insert Shift:=xlDown
            pengujiWS.Range(pengujiResultLocation1 & pengujiLastRow).SpecialCells(xlCellTypeVisible).Copy MainWS.Range("H" & pengujiFirstRowDestination)
            pengujiWS.Range(pengujiResultLocation2 & pengujiLastRow).SpecialCells(xlCellTypeVisible).Copy MainWS.Range("I" & pengujiFirstRowDestination)
            pengujiWS.Range(pengujiResultLocation3 & pengujiLastRow).SpecialCells(xlCellTypeVisible).Copy MainWS.Range("J" & pengujiFirstRowDestination)
            pengujiWS.Range(pengujiResultLocation4 & pengujiLastRow).SpecialCells(xlCellTypeVisible).Copy MainWS.Range("K" & pengujiFirstRowDestination)
            pengujiWS.Range(pengujiResultLocation5 & pengujiLastRow).SpecialCells(xlCellTypeVisible).Copy MainWS.Range("L" & pengujiFirstRowDestination)
            
    End If

End With

'===========================================================================================
'On Error Resume Next
Dim pemicuWS As Worksheet
Dim pemicuCopyRng As Range
Dim pemicuLastRow As Long
Dim pemicuCriteriaColumnNo_1 As Integer
Dim pemicuCriteriaColumnNo_2 As Integer

Dim pemicuResultLocation1 As String
Dim pemicuResultLocation2 As String
Dim pemicuResultLocation3 As String
Dim pemicuResultLocation4 As String
Dim pemicuResultLocation5 As String

pemicuCriteriaColumnNo_1 = 3
pemicuCriteriaColumnNo_2 = 15

'======================================= Destination ==================================
pemicuFirstRowDestination = 29

pemicuResultLocation1 = ResultColumn1 & firstTabelToBeFilterRow + 1 & ":" & ResultColumn1
pemicuResultLocation2 = ResultColumn2 & firstTabelToBeFilterRow + 1 & ":" & ResultColumn2
pemicuResultLocation3 = ResultColumn3 & firstTabelToBeFilterRow + 1 & ":" & ResultColumn3
pemicuResultLocation4 = ResultColumn4 & firstTabelToBeFilterRow + 1 & ":" & ResultColumn4
pemicuResultLocation5 = ResultColumn5 & firstTabelToBeFilterRow + 1 & ":" & ResultColumn5

Set pemicuWS = Sheets("DATA PEMICU")

pemicuLastRow = pemicuWS.Cells.Find("*", SearchOrder:=xlByRows, _
SearchDirection:=xlPrevious).Row


'============================Copy Data Pemicu to other Sheet to remove Formula=============================

pemicuWS.Range("W15:W" & pemicuLastRow).Copy
pemicuWS.Range("X15:X" & pemicuLastRow).PasteSpecial Paste:=xlPasteValues
pemicuWS.Columns("X:X").NumberFormat = _
        "_([$Rp-id-ID]* #,##0_);_([$Rp-id-ID]* (#,##0);_([$Rp-id-ID]* ""-""_);_(@_)"

'============================Filtering and Copy Process Pemicu=============================

With pemicuWS.Range("G15:X" & pemicuLastRow)
    On Error Resume Next
    .AutoFilter field:=pemicuCriteriaColumnNo_1, Criteria1:=AccountUsed
    .AutoFilter field:=pemicuCriteriaColumnNo_2, Criteria1:="Digunakan"
    If pemicuWS.Range(pemicuResultLocation1 & pemicuLastRow).SpecialCells(xlCellTypeVisible).Cells.Count = 0 Then
            MainWS.Range("H" & pemicuFirstRowDestination) = "Data Pemicu Tidak Tersedia"
            
    ElseIf pemicuWS.Range(pemicuResultLocation1 & pemicuLastRow).SpecialCells(xlCellTypeVisible).Cells.Count = 1 Then
            pemicuWS.Range(pemicuResultLocation1 & pemicuLastRow).SpecialCells(xlCellTypeVisible).Copy MainWS.Range("H" & pemicuFirstRowDestination)
            pemicuWS.Range(pemicuResultLocation2 & pemicuLastRow).SpecialCells(xlCellTypeVisible).Copy MainWS.Range("I" & pemicuFirstRowDestination)
            pemicuWS.Range(pemicuResultLocation3 & pemicuLastRow).SpecialCells(xlCellTypeVisible).Copy MainWS.Range("J" & pemicuFirstRowDestination)
            pemicuWS.Range(pemicuResultLocation4 & pemicuLastRow).SpecialCells(xlCellTypeVisible).Copy MainWS.Range("K" & pemicuFirstRowDestination)
            pemicuWS.Range(pemicuResultLocation5 & pemicuLastRow).SpecialCells(xlCellTypeVisible).Copy MainWS.Range("L" & pemicuFirstRowDestination)
            
            
    ElseIf pemicuWS.Range(pemicuResultLocation1 & pemicuLastRow).SpecialCells(xlCellTypeVisible).Cells.Count > 1 Then
            MainWS.Range(pemicuFirstRowDestination + 1 & ":" & pemicuWS.AutoFilter.Range.Columns(firstTabelToBeFilterRow).SpecialCells(xlCellTypeVisible).Cells.Count + pemicuFirstRowDestination - 1).Insert Shift:=xlDown
            pemicuWS.Range(pemicuResultLocation1 & pemicuLastRow).SpecialCells(xlCellTypeVisible).Copy MainWS.Range("H" & pemicuFirstRowDestination)
            pemicuWS.Range(pemicuResultLocation2 & pemicuLastRow).SpecialCells(xlCellTypeVisible).Copy MainWS.Range("I" & pemicuFirstRowDestination)
            pemicuWS.Range(pemicuResultLocation3 & pemicuLastRow).SpecialCells(xlCellTypeVisible).Copy MainWS.Range("J" & pemicuFirstRowDestination)
            pemicuWS.Range(pemicuResultLocation4 & pemicuLastRow).SpecialCells(xlCellTypeVisible).Copy MainWS.Range("K" & pemicuFirstRowDestination)
            pemicuWS.Range(pemicuResultLocation5 & pemicuLastRow).SpecialCells(xlCellTypeVisible).Copy MainWS.Range("L" & pemicuFirstRowDestination)
            
   
    End If

End With

'=======================================Formating=======================================

    Range("G" & pemicuFirstRowDestination - 1).CurrentRegion.Select
    Selection.Borders(xlDiagonalDown).LineStyle = xlNone
    Selection.Borders(xlDiagonalUp).LineStyle = xlNone
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .ColorIndex = 0
        .TintAndShade = 0
        .Weight = xlThin
    End With

Call LockEWS.lockPemicu
Call LockEWS.lockPenguji
MainWS.Select
Range("A1").Select
End Sub