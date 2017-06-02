Sub Splt()
Dim LR As Long, i As Long
Dim X As Variant
Application.ScreenUpdating = False
LR = Range("G" & Rows.Count).End(xlUp).Row
Columns("G").Insert
For i = LR To 1 Step -1
    With Range("H" & i)
        If InStr(.Value, "<DIV>") = 0 Then
            .Offset(, -1).Value = .Value
        Else
            X = Split(.Value, "<DIV>")
            .Offset(1).Resize(UBound(X)).EntireRow.Insert
            .Offset(, -1).Resize(UBound(X) - LBound(X) + 1).Value = Application.Transpose(X)
        End If
    End With
Next i
Columns()(8).EntireColumn.Delete
LR = Range("G" & Rows.Count).End(xlUp).Row
With Range("A3:Q" & LR)
    On Error Resume Next
    .SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[-1]C"
    On Error GoTo 0
    .Value = .Value
End With
Rows()(3).EntireRow.Delete
Rows()(1).EntireRow.Delete
Columns()(8).EntireColumn.Delete
Application.ScreenUpdating = True
End Sub

