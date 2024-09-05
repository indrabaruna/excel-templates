Sub PengujiTableFilter()

Call LockEWS.unlockPenguji
Dim AKUN, NAMA_DP, DW_SK_PENGUJI_H, DW_START_TIME, NPWP, KD_PENGUJI, ID_MS_TH_PJK, DATA_JSON, STATUS, KETERANGAN As String
Dim NILAI1, NILAI2, NILAI_DATA, NILAI_YANG_DIGUNAKAN As Long
Dim lastRow As Long
Dim MainWS As Worksheet
Set MainWS = Worksheets("DATA PENGUJI")


With MainWS
lastRow = .Cells.Find("*", SearchOrder:=xlByRows, _
SearchDirection:=xlPrevious).Row
If lastRow < 15 Then lastRow = 15
If .Range("H6").Value = "Ketik Akun" Then AKUN = Empty Else: AKUN = .Range("H6").Value
.Range("G15:W" & lastRow).Select
Selection.AutoFilter
    With .Range("G15:W" & lastRow)
        If AKUN <> Empty Then .AutoFilter field:=3, Criteria1:="=*" & AKUN & "*"
        End With
        
.Range("15:15").EntireRow.Hidden = True

End With
Call LockEWS.lockPenguji
MainWS.Select
Range("A1").Select

End Sub

Sub EmptyPengujiTableFilter()
Call LockEWS.unlockPenguji
Dim AKUN, NAMA_DP, DW_SK_PENGUJI_H, DW_START_TIME, NPWP, KD_PENGUJI, ID_MS_TH_PJK, DATA_JSON, STATUS, KETERANGAN As String
Dim NILAI1, NILAI2, NILAI_DATA, NILAI_YANG_DIGUNAKAN As Long
Dim lastRow As Long
Dim MainWS As Worksheet
Set MainWS = Worksheets("DATA PENGUJI")

With MainWS
lastRow = .Cells.Find("*", SearchOrder:=xlByRows, _
SearchDirection:=xlPrevious).Row
.Range("G15:W" & lastRow).Select
Selection.AutoFilter
    With .Range("G15:W" & lastRow)
    .AutoFilter field:=3, Criteria1:=""
        End With
        
.Range("15:15").EntireRow.Hidden = True

End With
Call LockEWS.lockPenguji
MainWS.Select
Range("A1").Select

End Sub

Sub PengujiClearFilter()

'Call lockEWS.unlockPenguji
Dim MainWS As Worksheet
Set MainWS = Worksheets("DATA PENGUJI")

With MainWS
.AutoFilterMode = False
.Range("H6").Value = "Ketik Akun"
.Range("H10").Value = "Tidak ada ID Data yang Dipilih"

End With
'Call lockEWS.lockPenguji
End Sub
'Sub FilteringDetailPenguji()
'
'Dim DW_SK_PENGUJI_D, DW_SK_PENGUJI_H, KD_PENGUJI, DATA_JSON As String
'Dim lastRow As Long
'Dim MainWS As Worksheet
'Dim PENGUJIMainWS As Worksheet
'Set MainWS = Worksheets("DPENGUJI DETAIL")
'Set PENGUJIMainWS = Worksheets("DATA PENGUJI")
'
'lastRow = MainWS.Cells.Find("*", SearchOrder:=xlByRows, _
'SearchDirection:=xlPrevious).Row
'If lastRow < 15 Then lastRow = 15
'        If PENGUJIMainWS.Range("H10").Value = "Tidak ada ID Data yang Dipilih" Then DW_SK_PENGUJI_H = Empty Else: DW_SK_PENGUJI_H = PENGUJIMainWS.Range("H10").Value
'MainWS.Activate
'MainWS.Range("G15:J" & lastRow).Select
'Selection.AutoFilter
'    With MainWS.Range("G15:J" & lastRow)
'        If DW_SK_PENGUJI_H <> Empty Then .AutoFilter field:=3, Criteria1:=PENGUJIMainWS.Range("H10").Value
'
'MainWS.Range("15:15").EntireRow.Hidden = True
'
'End With
'End Sub

Sub FilteringDetailPenguji()
Call LockEWS.unlockPenguji
Call LockEWS.unlockDetailPenguji

Dim DW_SK_PENGUJI_D, DW_SK_PENGUJI_H, KD_PENGUJI, DATA_JSON As String
Dim lastRow As Long
Dim MainWS As Worksheet
Dim pengujiMainWS As Worksheet
Set MainWS = Worksheets("DPENGUJI DETAIL")
Set pengujiMainWS = Worksheets("DATA PENGUJI")


lastRow = MainWS.Cells.Find("*", SearchOrder:=xlByRows, _
SearchDirection:=xlPrevious).Row
If lastRow < 15 Then lastRow = 15
If pengujiMainWS.Range("H10").Value = "Tidak ada ID Data yang Dipilih" Then DW_SK_PENGUJI_H = Empty Else: DW_SK_PENGUJI_H = pengujiMainWS.Range("H10").Value
MainWS.Activate
MainWS.Range("G15:J" & lastRow).Select
Selection.AutoFilter
    With MainWS.Range("G15:J" & lastRow)
        If DW_SK_PENGUJI_H <> Empty Then .AutoFilter field:=2, Criteria1:=pengujiMainWS.Range("H10").Value

MainWS.Range("15:15").EntireRow.Hidden = True

End With
Call LockEWS.lockPenguji
Call LockEWS.lockDetailPenguji
MainWS.Select
Range("A1").Select

End Sub
Sub ClearPengujiFilterDetail()
Call LockEWS.unlockDetailPenguji
Dim MainWS As Worksheet
Set MainWS = Worksheets("DPENGUJI DETAIL")

With MainWS
.AutoFilterMode = False
End With
Call LockEWS.lockDetailPenguji
MainWS.Select
Range("A1").Select
End Sub