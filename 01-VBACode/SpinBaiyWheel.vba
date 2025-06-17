Sub SpinBaiyWheel()
    
    Dim baiyRewardArr(1 to 12) As Double
    Dim baiyProbabilityArr(1 to 12) As Double
    Dim result As Double

    Dim i As Long
    
    For i=1 To 12
        baiyRewardArr(i) = ThisWorkbook.Sheets("转盘").Range("E3").Offset(0, i - 1).Value
        baiyProbabilityArr(i) = ThisWorkbook.Sheets("转盘").Range("E4").Offset(0, i - 1).Value
    Next i

    'For i = 1 To 12
    '    Debug.Print "baiyRewardArr(" & i & ") = " & baiyRewardArr(i)
    'Next i

    'For i = 1 To 12
    '    Debug.Print "baiyProbabilityArr(" & i & ") = " & baiyProbabilityArr(i)
    'Next i

    
    '0-1随机一个数
    Randomize
    randNum = Rnd
    'Debug.Print "randNum = " & randNum '打印随机数

    '使用累计概率判断
    Dim cumProb As Double
    cumProb=0
    For i = 1 to 12
        cumProb=cumProb+baiyProbabilityArr(i)
        If  randNum<=cumProb Then
            result=baiyRewardArr(i)
            'ThisWorkbook.Sheets("转盘").Range("T1").Offset(1, 0).Value="您这次白银转盘获得"&CStr(result)&"个小票"
            Exit For
        End If
    Next i

    ' 找到第一个空白的 T列位置，从 T2 开始往下找
    Dim tRow As Long
    tRow = 2 ' T2起始行
    Do While Not IsEmpty(ThisWorkbook.Sheets("转盘").Range("T" & tRow))
        tRow = tRow + 1
        ' 若想限制最大行数，可以加个条件，例如到T100
        If tRow > 1000 Then Exit Sub
    Loop
    
    ' 在找到的空白行写入
    ThisWorkbook.Sheets("转盘").Range("T" & tRow).Value = "您这次白银转盘获得 " & CStr(result) & " 个小票"

End Sub
