Attribute VB_Name = "matrix"

Sub main()
    
    Call init
    
    '�V�K�V�[�g�̒ǉ�
    'Worksheets.Add after:=Worksheets("����")
    
    


End Sub


Function init()
    
    '�A�N�e�B�u�V�[�g�̕ύX
    Worksheets("����").Activate

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


'�w�肳�ꂽ�J�����̍ő�s���擾����
Function fncGetRowSize(column As Long) As Long
    Dim row As Long
    row = 1
    Do While Cells(row + 1, column).Value <> ""
        row = row + 1
    Loop
    fncGetRowSize = row
End Function
