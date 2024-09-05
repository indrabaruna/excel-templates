Sub dataValidationPengujian_B1()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 48
rowPemilihankesimpulan = 58
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("B-1-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_B2()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 48
rowPemilihankesimpulan = 58
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("B-2-1")
    
    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_B3()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("B-3-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_B4()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("B-4-1")
    
    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_B5()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("B-5-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_C1()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("C-1-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_C2()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("C-2-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub
Sub dataValidationPengujian_C3()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("C-3-1")
'

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_C4()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("C-4-1")
'

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_C5()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("C-5-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_C6()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("C-6-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_C7()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("C-7-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_C8()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 34
rowPemilihankesimpulan = 44
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("C-8-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_C9()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("C-9-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_C10()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("C-10-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_C11()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("C-11-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_C12()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("C-12-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_C13()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("C-13-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_C14()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("C-14-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_D()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("D-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_E1()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("E-1-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_E2()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("E-2-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_E3()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("E-3-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_E4()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("E-4-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_E5()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("E-5-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_E6()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("E-6-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_E7()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("E-7-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_E8()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("E-8-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_E9()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("E-9-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_E10()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("E-10-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_E11()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("E-11-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_E12()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("E-12-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_E13()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("E-13-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    
    End With

End Sub


Sub dataValidationPengujian_F1()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("F-1-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_F2()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 34
rowPemilihankesimpulan = 44
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("F-2-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_F3()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("F-3-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_F4()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("F-4-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_G1()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("G-1-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_G2()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("G-2-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_G3()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("G-3-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_G4()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 34
rowPemilihankesimpulan = 44
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("G-4-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_H()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 34
rowPemilihankesimpulan = 44
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("H-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_I()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 34
rowPemilihankesimpulan = 44
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("I-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_J()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 34
rowPemilihankesimpulan = 44
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("J-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_K()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 38
rowPemilihankesimpulan = 48
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("K-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_L()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 34
rowPemilihankesimpulan = 44
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("L-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_M1()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 34
rowPemilihankesimpulan = 44
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("M-1-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_M2()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 34
rowPemilihankesimpulan = 44
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("M-2-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_M3()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 34
rowPemilihankesimpulan = 44
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("M-3-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_P()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 34
rowPemilihankesimpulan = 44
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("P-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujian_Q()

Dim WS As Worksheet
Dim rowPemilihanTeknik As Integer
Dim firstRowResult As Integer
Dim firstRowPerkiraan As Integer
Dim lRow As Integer
Dim rangeKesimpulan As String

firstRowResult = 16
firstRowPerkiraan = 16
rowPemilihanTeknik = 34
rowPemilihankesimpulan = 44
lRow = Cells(Rows.Count, 1).End(xlUp).Row
rangeKesimpulan = "='MASTER'!$Q$2:$Q$6"
Set WS = Sheets("Q-1")

    WS.Activate
    WS.Range("A" & firstRowResult).Select
    ActiveCell.Formula2R1C1 = "=FILTER(RC[12]:R[34]C[12],RC[12]:R[34]C[12]<>"""")"
    WS.Range("I" & rowPemilihanTeknik & ":V" & rowPemilihanTeknik).Select
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:="=$A$" & firstRowPerkiraan & "#"
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

    WS.Range("I" & rowPemilihankesimpulan & ":AB" & rowPemilihankesimpulan).Select
    
    
'    Application.CutCopyMode = False
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:= _
        xlBetween, Formula1:=rangeKesimpulan
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With

End Sub


Sub dataValidationPengujianAll()

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

End Sub