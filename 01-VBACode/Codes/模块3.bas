Attribute VB_Name = "模块3"
Sub OpenFile()
    If Application.FindFile = True Then
        Debug.Print ("已打开")
    Else
        Debug.Print ("未打开")
    End If
    
'    Application.FindFile
End Sub


Sub 案例InputBox()
    Dim str As String
    str = InputBox("请输入你的名字")
    Range("e112").Value = str
    MsgBox ("你成功输入了你的名字"), Buttons:=48
End Sub


Sub 案例With()
    With Range("A120").Font
        .ColorIndex = 3
        .Size = 13
        .Bold = True
    End With
End Sub


Sub 案例计数2()
    Dim i As Integer
    xrow = 85
    For i = 3 To 100 Step 3
        Cells(xrow, "A") = i
        xrow = xrow + 1
    Next i
End Sub


Sub SpinHuangjWheel()
    
    Dim huangjRewardArr(1 To 12) As Double
    Dim huangjProbabilityArr(1 To 12) As Double
    Dim result As Double

    Dim i As Long
    
    For i = 1 To 12
        huangjRewardArr(i) = ThisWorkbook.Sheets("转盘").Range("E9").Offset(0, i - 1).Value
        huangjProbabilityArr(i) = ThisWorkbook.Sheets("转盘").Range("E10").Offset(0, i - 1).Value
    Next i

    '0-1随机一个数
    Randomize
    randNum = Rnd

    '使用累计概率判断
    Dim cumProb As Double
    cumProb = 0
    For i = 1 To 12
        cumProb = cumProb + huangjProbabilityArr(i)
        If randNum <= cumProb Then
            result = huangjRewardArr(i)
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
    ThisWorkbook.Sheets("转盘").Range("T" & tRow).Value = "您这次黄金转盘获得 " & CStr(result) & " 个小票"

End Sub
Sub SpinZuansWheel()
    
    Dim zuansRewardArr(1 To 12) As Double
    Dim zuansProbabilityArr(1 To 12) As Double
    Dim result As Double

    Dim i As Long
    
    For i = 1 To 12
        zuansRewardArr(i) = ThisWorkbook.Sheets("转盘").Range("E15").Offset(0, i - 1).Value
        zuansProbabilityArr(i) = ThisWorkbook.Sheets("转盘").Range("E16").Offset(0, i - 1).Value
    Next i

    '0-1随机一个数
    Randomize
    randNum = Rnd

    '使用累计概率判断
    Dim cumProb As Double
    cumProb = 0
    For i = 1 To 12
        cumProb = cumProb + zuansProbabilityArr(i)
        If randNum <= cumProb Then
            result = zuansRewardArr(i)
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
    ThisWorkbook.Sheets("转盘").Range("T" & tRow).Value = "您这次钻石转盘获得 " & CStr(result) & " 个小票"

End Sub

Sub DeleteSpinWheelLog()
    ThisWorkbook.Sheets("转盘").Range("T2:T1000").Delete
End Sub
Sub SpinWheelWithProbabilities()
    Dim results As Range
    Dim probabilities As Range
    Dim totalProb As Double
    Dim randNum As Double
    Dim cumulativeProb As Double
    Dim i As Long
    Dim logSheet As Worksheet
    Dim logRow As Long
    
    ' 设置数据范围
    Set results = ThisWorkbook.Sheets("转盘").Range("E3:P3")
    Set probabilities = ThisWorkbook.Sheets("转盘").Range("E4:P4")
    
    ' 计算概率总和，确保为1
    totalProb = Application.WorksheetFunction.Sum(probabilities)
    If totalProb <= 0 Then
        MsgBox "概率总和为0或负数，请检查概率值是否设置正确。"
        Exit Sub
    End If
    
    ' 生成0到1之间的随机数
    Randomize
    randNum = Rnd
    
    ' 通过概率区间找到对应的结果
    cumulativeProb = 0
    Dim selectedResult As String
    
    For i = 1 To 12
        cumulativeProb = cumulativeProb + probabilities.Cells(i, 1).Value / totalProb
        If randNum <= cumulativeProb Then
            selectedResult = results.Cells(i, 1).Value
            Exit For
        End If
    Next i
    
    ' 在B1显示结果
    ThisWorkbook.Sheets("转盘").Range("R1").Value = "结果: " & Format(selectedResult, "0.00")
    
    ' 记录到日志
    Set logSheet = ThisWorkbook.Sheets("转盘")
    logRow = logSheet.Cells(logSheet.Rows.Count, "S").End(xlUp).Row + 1
    If logSheet.Cells(logRow, "S").Value <> "" Then
        logRow = logRow + 1
    End If
    logSheet.Cells(logRow, "S").Value = "旋转时间: " & Now
    logSheet.Cells(logRow, "T").Value = "结果是: " & selectedResult
End Sub
Sub SpinBaiyWheel()
    
    Dim baiyRewardArr(1 To 12) As Double
    Dim baiyProbabilityArr(1 To 12) As Double
    Dim result As Double

    Dim i As Long
    
    For i = 1 To 12
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
    cumProb = 0
    For i = 1 To 12
        cumProb = cumProb + baiyProbabilityArr(i)
        If randNum <= cumProb Then
            result = baiyRewardArr(i)
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

Sub 案例Select()
    Dim cj As Variant
    cj = InputBox("输入考试成绩：")
    Select Case cj
    Case 0 To 59
        MsgBox "等级：D"
    Case 60 To 69
        MsgBox "等级：C"
    Case 70 To 79
        MsgBox "等级：B"
    Case 80 To 100
        MsgBox "等级：A"
    Case Else
        MsgBox "输入错误"
    End Select
End Sub

Sub 案例计数()
    Dim i As Integer
    xrow = 30
    For i = 1 To 100 Step 2
        Cells(xrow, "A") = i
        xrow = xrow + 1
    Next i
End Sub

Sub 查名字()
    If Range("a18") Like "邱*" Then
    Range("A19").Value = "ta姓邱"
    Else
    Range("A19").Value = "ta不姓邱"
    End If
End Sub

Sub 自定义旋转转盘()
 旋转转盘框体.Show
End Sub
