Sub SpinWheelWithProbabilities()
    Dim results As Range
    Dim probabilities As Range
    Dim totalProb As Double
    Dim randNum As Double
    Dim cumulativeProb As Double
    Dim i As long
    Dim logSheet As Worksheet
    Dim logRow As Long
    
    ' 设置数据范围
    Set results = ThisWorkbook.Sheets("转盘").Range("E3:P3")
    Set probabilities = ThisWorkbook.Sheets("转盘").Range("E4:P4")
    
    ' 计算概率总和，确保为1
    totalProb = Application.WorksheetFunction.Sum(probabilities)
    If totalProb <= 0 Then
        MsgBox "概率总和为0或负数,请检查概率值是否设置正确。"
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
    logSheet.Cells(logRow, "T").Value = "结果: " & selectedResult
End Sub
