```vb
Sub MPS()
    Dim value01 As Integer
    '定义行列变量
    Dim row As Integer
    Dim column As Integer
    row = Selection.row
    column = Selection.column
    '定义循环变量
    Dim t As Integer
    t = 1
    '定义最大周期数
    Dim tMax As Integer
    tMax = InputBox("输入总周期数:")
    '定义数组为100的长度
    Dim I(99)
    I(0) = InputBox("输入期初库存:")
    Dim M As Integer
    M = InputBox("输入计划生产量:")
    While (t <= tMax)
        If (Cells(row - 2, column) >= Cells(row - 1, column)) Then
            I(t) = I(t - 1) - Cells(row - 2, column)
            Cells(row + 1, column) = 0
        Else
            I(t) = I(t - 1) - Cells(row - 1, column)
            Cells(row + 1, column) = 0
        End If
        If (I(t) < 0) Then
            I(t) = I(t) + M
            Cells(row + 1, column) = M
        End If
        Cells(row, column) = I(t)
        t = t + 1
        column = column + 1
    Wend
    
    row = Selection.row
    column = Selection.column
    t = 1
    Dim j As Integer
    j = 0
    Dim A As Integer
    A = I(0)
    
    While (Not (Cells(row + 1, column + j) <> 0))
        A = A - Cells(row - 1, column + j)
        j = j + 1
    Wend
    Cells(row + 2, column) = A
    t = column + j
    j = 1
    
    While (t <= tMax)
        If (Cells(row + 1, t) <> 0) Then
            A = M - Cells(row - 1, t)
            While (Not (Cells(row + 1, t + j) <> 0))
                A = A - Cells(row - 1, t + j)
                j = j + 1
            Wend
            Cells(row + 2, t) = A
            j = 0
        End If
        t = t + 1
    Wend
End Sub

```

