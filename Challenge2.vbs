Sub Mod2():
    Dim I, j As Integer
    Dim ticker As String
    Dim totalstock As Double
    totalstock = 0
    Dim Summary_Table_Row As Integer
    Summary_Table_Row = 2
    n = Worksheets("A").UsedRange.Rows.Count

        For Each Current In Worksheets

            For I = 2 To n
                If Cells(I + 1, 1).Value <> Cells(I, 1).Value Then
                    ticker = Cells(I, 1).Value
                    totalstock = totalstock + Cells(I, 7).Value
                    Range("L" & Summary_Table_Row).Value = totalstock
                    Range("I" & Summary_Table_Row).Value = ticker
                    Summary_Table_Row = Summary_Table_Row + 1
                    totalstock = 0
                Else
                    totalstock = totalstock + Cells(I, 7).Value
                End If
             Next I
            MsgBox Current.Name
         Next
End Sub
