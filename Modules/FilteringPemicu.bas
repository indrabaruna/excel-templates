Option Explicit

Sub PemicuTableFilter()

Call LockEWS.unlockPemicu
Dim JENIS, AKUN, NAMA_DP, DW_SK_PEMICU_H, DW_START_TIME, NPWP, KD_PEMICU, ID_MS_TH_PJK, ID_JNS_WP, DATA_JSON, STATUS, KETERANGAN As String
Dim NILAI1, NILAI2, NILAI_DATA, NILAI_YANG_DIGUNAKAN As Long
Dim lastRow As Long
Dim MainWS As Worksheet
Set MainWS = Worksheets("DATA PEMICU")


With MainWS
lastRow = .Cells.Find("*", SearchOrder:=xlByRows, _
SearchDirection:=xlPrevious).Row
If lastRow < 15 Then lastRow = 15
If .Range("H7").Value = "Ketik Akun" Then AKUN = Empty Else: AKUN = .Range("H7").Value
.Range("G15:W" & lastRow).Select
Selection.AutoFilter
    With .Range("G15:W" & lastRow)
        If AKUN <> Empty Then .AutoFilter field:=3, Criteria1:="=*" & AKUN & "*"
        End With
        
.Range("15:15").EntireRow.Hidden = True

End With
Call LockEWS.lockPemicu
MainWS.Select
Range("A1").Select
End Sub

Sub PemicuClearFilter()

'Call lockEWS.unlockPemicu
Dim MainWS As Worksheet
Set MainWS = Worksheets("DATA PEMICU")

With MainWS
.AutoFilterMode = False
.Range("H7").Value = "Ketik Akun"
.Range("H10").Value = "Tidak ada ID Data yang Dipilih"

End With
'Call lockEWS.lockPemicu
End Sub

Sub FilteringDetailPemicu()

Call LockEWS.unlockPemicu
Call LockEWS.unlockDetailPemicu

Dim DW_SK_PEMICU_D, DW_SK_PEMICU_H, KD_PEMICU, DATA_JSON As String
Dim lastRow As Long
Dim MainWS As Worksheet
Dim PemicuMainWS As Worksheet
Set MainWS = Worksheets("DPEMICU DETAIL")
Set PemicuMainWS = Worksheets("DATA PEMICU")


lastRow = MainWS.Cells.Find("*", SearchOrder:=xlByRows, _
SearchDirection:=xlPrevious).Row
If lastRow < 15 Then lastRow = 15
If PemicuMainWS.Range("H10").Value = "Tidak ada ID Data yang Dipilih" Then DW_SK_PEMICU_H = Empty Else: DW_SK_PEMICU_H = PemicuMainWS.Range("H10").Value
MainWS.Activate
MainWS.Range("G15:J" & lastRow).Select
Selection.AutoFilter
    With MainWS.Range("G15:J" & lastRow)
        If DW_SK_PEMICU_H <> Empty Then .AutoFilter field:=2, Criteria1:=PemicuMainWS.Range("H10").Value

MainWS.Range("15:15").EntireRow.Hidden = True

End With
Call LockEWS.lockPemicu
Call LockEWS.lockDetailPemicu
MainWS.Select
Range("A1").Select
End Sub

Sub ClearPemicuFilterDetail()

Call LockEWS.unlockDetailPemicu
Dim MainWS As Worksheet
Set MainWS = Worksheets("DPEMICU DETAIL")

With MainWS
.AutoFilterMode = False
End With
Call LockEWS.lockDetailPemicu
MainWS.Select
Range("A1").Select

End Sub