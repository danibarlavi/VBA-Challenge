Attribute VB_Name = "Module2"
Sub Conditional_Formatting()
    For Each ws In Worksheets
       'Colors all positive yearly changes green and all negative yearly changes red
        For i = 2 To 2977
            If ws.Range("J" & i).Value > 0 Then
                ws.Range("J" & i).Interior.ColorIndex = 4
            Else
                ws.Range("J" & i).Interior.ColorIndex = 3
            End If
        Next i
    Next ws
End Sub
