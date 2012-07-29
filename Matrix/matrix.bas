Attribute VB_Name = "matrix"

Dim OUTPUT_SHEET_NAME As String

Public Sub PROC_MATRIX_CREATER()
    Call init
    Call main
    Call finally
End Sub


Private Function init()
    Call fncCreateOutputSheet
End Function

Private Function main()
    Call fncLoop(1, 1)
End Function

Private Function finally()
    Worksheets(OUTPUT_SHEET_NAME).Activate
End Function

Private Function fncCreateOutputSheet()
    
    Dim active_sheet_name As String
    active_sheet_name = activeSheet.name

    Worksheets.Add after:=Worksheets(active_sheet_name)
    OUTPUT_SHEET_NAME = activeSheet.name

    Worksheets(active_sheet_name).Activate

End Function


Private Function fncLoop(row As Long, col As Long)
    
    Dim rtnVal As Boolean
    Dim rowsize As Long
    rowsize = fncGetRowSize(col)
    
    For i = 1 To rowsize
        
        Worksheets(OUTPUT_SHEET_NAME).Cells(row, col).Value = Cells(i, col).Value
        
        If Cells(1, col + 1).Value <> "" Then
            Call fncLoop(row, col + 1)
        Else
            row = row + 1
        End If
    
    Next

End Function

Private Function fncGetRowSize(col As Long) As Long
    Dim row As Long
    row = 1
    Do While Cells(row + 1, col).Value <> ""
        row = row + 1
    Loop
    fncGetRowSize = row
End Function
