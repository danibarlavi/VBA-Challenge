Attribute VB_Name = "Module3"

Sub Max_Finder()
    For Each ws In Worksheets
        'Find the Greatest Increase, Decrease, and Volume
        Dim GreatestIncrease As Double
        GreatestIncrease = WorksheetFunction.Max(ws.Range("K2:K2977"))
        Dim GreatestDecrease As Double
        GreatestDecrease = WorksheetFunction.Min(ws.Range("K2:K2977"))
        Dim GreatestVolume As Double
        GreatestVolume = WorksheetFunction.Max(ws.Range("L2:L2977"))
        
        'Print these values onto the summary table
        ws.Range("Q2") = GreatestIncrease
        ws.Range("Q3") = GreatestDecrease
        ws.Range("Q4") = GreatestVolume
        
        'Find the ticker corresponding to the Greatest Increase, Decrease, and Volume and print on the summary table. Shout out to Rocky Myers for help with this code.
        For i = 2 To 2977
            If ws.Cells(i, 11).Value = GreatestIncrease Then
            ws.Cells(2, 16).Value = ws.Cells(i, 9).Value
            ElseIf ws.Cells(i, 11).Value = GreatestDecrease Then
            ws.Cells(3, 16).Value = ws.Cells(i, 9).Value
            ElseIf ws.Cells(i, 12).Value = GreatestVolume Then
            ws.Cells(4, 16).Value = ws.Cells(i, 9).Value
            End If
        Next i
    Next ws
End Sub
