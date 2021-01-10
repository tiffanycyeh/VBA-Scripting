Sub Conditional_Format()

Dim ws As Worksheet
For Each ws In Worksheets

NumRows = Range("K2", Range("K2").End(xlDown)).Rows.Count

    For i = 2 To NumRows
    
        If ws.Cells(i, 11) >= 0 Then
            ws.Cells(i, 11).Interior.ColorIndex = 4
        Else: ws.Cells(i, 11).Interior.ColorIndex = 3
        End If
    Next i
Next ws

End Sub
