Attribute VB_Name = "matrix"

Sub main()
    
    Call init
    
    '新規シートの追加
    'Worksheets.Add after:=Worksheets("項目")
    
    


End Sub


Function init()
    
    'アクティブシートの変更
    Worksheets("項目").Activate

    Call fncLoop(1, 1)
    

End Function

Function fncLoop(row As Long, column As Long) As Boolean
    
    Dim rtnVal As Boolean
    Dim rowSize As Long
    rowSize = fncGetRowSize(column)
    
    For i = 1 To rowSize
        
        Cells(5 + row, column).Value = Cells(i, column).Value
        
        If Cells(i, column + 1).Value <> "" Then
            rtnVal = fncLoop(row, column + 1)
        End If
        
        If rtnVal <> True Then
            row = row + 1
        End If
    
    Next

    fncLoop = True
    

End Function


'指定されたカラムの最大行を取得する
Function fncGetRowSize(column As Long) As Long
    Dim row As Long
    row = 1
    Do While Cells(row + 1, column).Value <> ""
        row = row + 1
    Loop
    fncGetRowSize = row
End Function
