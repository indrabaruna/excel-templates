Sub ShowHiddenNames()

    '    Dimension variables
       Dim xName As Variant
       Dim Result As Variant
       Dim Vis As Variant

    '    Loop once for each name in the workbook
       For Each xName In ActiveWorkbook.Names

        '    If a name is not visible then make it visible
           If xName.Visible = False Then
               xName.Visible = True
           End If
       Next xName

End Sub