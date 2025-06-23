Attribute VB_Name = "ģ��3"
Sub OpenFile()
    If Application.FindFile = True Then
        Debug.Print ("�Ѵ�")
    Else
        Debug.Print ("δ��")
    End If
    
'    Application.FindFile
End Sub


Sub ����InputBox()
    Dim str As String
    str = InputBox("�������������")
    Range("e112").Value = str
    MsgBox ("��ɹ��������������"), Buttons:=48
End Sub


Sub ����With()
    With Range("A120").Font
        .ColorIndex = 3
        .Size = 13
        .Bold = True
    End With
End Sub


Sub ��������2()
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
        huangjRewardArr(i) = ThisWorkbook.Sheets("ת��").Range("E9").Offset(0, i - 1).Value
        huangjProbabilityArr(i) = ThisWorkbook.Sheets("ת��").Range("E10").Offset(0, i - 1).Value
    Next i

    '0-1���һ����
    Randomize
    randNum = Rnd

    'ʹ���ۼƸ����ж�
    Dim cumProb As Double
    cumProb = 0
    For i = 1 To 12
        cumProb = cumProb + huangjProbabilityArr(i)
        If randNum <= cumProb Then
            result = huangjRewardArr(i)
            Exit For
        End If
    Next i

    ' �ҵ���һ���հ׵� T��λ�ã��� T2 ��ʼ������
    Dim tRow As Long
    tRow = 2 ' T2��ʼ��
    Do While Not IsEmpty(ThisWorkbook.Sheets("ת��").Range("T" & tRow))
        tRow = tRow + 1
        ' ��������������������ԼӸ����������絽T100
        If tRow > 1000 Then Exit Sub
    Loop
    
    ' ���ҵ��Ŀհ���д��
    ThisWorkbook.Sheets("ת��").Range("T" & tRow).Value = "����λƽ�ת�̻�� " & CStr(result) & " ��СƱ"

End Sub
Sub SpinZuansWheel()
    
    Dim zuansRewardArr(1 To 12) As Double
    Dim zuansProbabilityArr(1 To 12) As Double
    Dim result As Double

    Dim i As Long
    
    For i = 1 To 12
        zuansRewardArr(i) = ThisWorkbook.Sheets("ת��").Range("E15").Offset(0, i - 1).Value
        zuansProbabilityArr(i) = ThisWorkbook.Sheets("ת��").Range("E16").Offset(0, i - 1).Value
    Next i

    '0-1���һ����
    Randomize
    randNum = Rnd

    'ʹ���ۼƸ����ж�
    Dim cumProb As Double
    cumProb = 0
    For i = 1 To 12
        cumProb = cumProb + zuansProbabilityArr(i)
        If randNum <= cumProb Then
            result = zuansRewardArr(i)
            Exit For
        End If
    Next i

    ' �ҵ���һ���հ׵� T��λ�ã��� T2 ��ʼ������
    Dim tRow As Long
    tRow = 2 ' T2��ʼ��
    Do While Not IsEmpty(ThisWorkbook.Sheets("ת��").Range("T" & tRow))
        tRow = tRow + 1
        ' ��������������������ԼӸ����������絽T100
        If tRow > 1000 Then Exit Sub
    Loop
    
    ' ���ҵ��Ŀհ���д��
    ThisWorkbook.Sheets("ת��").Range("T" & tRow).Value = "�������ʯת�̻�� " & CStr(result) & " ��СƱ"

End Sub

Sub DeleteSpinWheelLog()
    ThisWorkbook.Sheets("ת��").Range("T2:T1000").Delete
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
    
    ' �������ݷ�Χ
    Set results = ThisWorkbook.Sheets("ת��").Range("E3:P3")
    Set probabilities = ThisWorkbook.Sheets("ת��").Range("E4:P4")
    
    ' ��������ܺͣ�ȷ��Ϊ1
    totalProb = Application.WorksheetFunction.Sum(probabilities)
    If totalProb <= 0 Then
        MsgBox "�����ܺ�Ϊ0�������������ֵ�Ƿ�������ȷ��"
        Exit Sub
    End If
    
    ' ����0��1֮��������
    Randomize
    randNum = Rnd
    
    ' ͨ�����������ҵ���Ӧ�Ľ��
    cumulativeProb = 0
    Dim selectedResult As String
    
    For i = 1 To 12
        cumulativeProb = cumulativeProb + probabilities.Cells(i, 1).Value / totalProb
        If randNum <= cumulativeProb Then
            selectedResult = results.Cells(i, 1).Value
            Exit For
        End If
    Next i
    
    ' ��B1��ʾ���
    ThisWorkbook.Sheets("ת��").Range("R1").Value = "���: " & Format(selectedResult, "0.00")
    
    ' ��¼����־
    Set logSheet = ThisWorkbook.Sheets("ת��")
    logRow = logSheet.Cells(logSheet.Rows.Count, "S").End(xlUp).Row + 1
    If logSheet.Cells(logRow, "S").Value <> "" Then
        logRow = logRow + 1
    End If
    logSheet.Cells(logRow, "S").Value = "��תʱ��: " & Now
    logSheet.Cells(logRow, "T").Value = "�����: " & selectedResult
End Sub
Sub SpinBaiyWheel()
    
    Dim baiyRewardArr(1 To 12) As Double
    Dim baiyProbabilityArr(1 To 12) As Double
    Dim result As Double

    Dim i As Long
    
    For i = 1 To 12
        baiyRewardArr(i) = ThisWorkbook.Sheets("ת��").Range("E3").Offset(0, i - 1).Value
        baiyProbabilityArr(i) = ThisWorkbook.Sheets("ת��").Range("E4").Offset(0, i - 1).Value
    Next i

    'For i = 1 To 12
    '    Debug.Print "baiyRewardArr(" & i & ") = " & baiyRewardArr(i)
    'Next i

    'For i = 1 To 12
    '    Debug.Print "baiyProbabilityArr(" & i & ") = " & baiyProbabilityArr(i)
    'Next i

    
    '0-1���һ����
    Randomize
    randNum = Rnd
    'Debug.Print "randNum = " & randNum '��ӡ�����

    'ʹ���ۼƸ����ж�
    Dim cumProb As Double
    cumProb = 0
    For i = 1 To 12
        cumProb = cumProb + baiyProbabilityArr(i)
        If randNum <= cumProb Then
            result = baiyRewardArr(i)
            'ThisWorkbook.Sheets("ת��").Range("T1").Offset(1, 0).Value="����ΰ���ת�̻��"&CStr(result)&"��СƱ"
            Exit For
        End If
    Next i

    ' �ҵ���һ���հ׵� T��λ�ã��� T2 ��ʼ������
    Dim tRow As Long
    tRow = 2 ' T2��ʼ��
    Do While Not IsEmpty(ThisWorkbook.Sheets("ת��").Range("T" & tRow))
        tRow = tRow + 1
        ' ��������������������ԼӸ����������絽T100
        If tRow > 1000 Then Exit Sub
    Loop
    
    ' ���ҵ��Ŀհ���д��
    ThisWorkbook.Sheets("ת��").Range("T" & tRow).Value = "����ΰ���ת�̻�� " & CStr(result) & " ��СƱ"

End Sub

Sub ����Select()
    Dim cj As Variant
    cj = InputBox("���뿼�Գɼ���")
    Select Case cj
    Case 0 To 59
        MsgBox "�ȼ���D"
    Case 60 To 69
        MsgBox "�ȼ���C"
    Case 70 To 79
        MsgBox "�ȼ���B"
    Case 80 To 100
        MsgBox "�ȼ���A"
    Case Else
        MsgBox "�������"
    End Select
End Sub

Sub ��������()
    Dim i As Integer
    xrow = 30
    For i = 1 To 100 Step 2
        Cells(xrow, "A") = i
        xrow = xrow + 1
    Next i
End Sub

Sub ������()
    If Range("a18") Like "��*" Then
    Range("A19").Value = "ta����"
    Else
    Range("A19").Value = "ta������"
    End If
End Sub

Sub �Զ�����תת��()
 ��תת�̿���.Show
End Sub
