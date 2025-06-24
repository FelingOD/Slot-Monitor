VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} 旋转转盘框体 
   Caption         =   "旋转转盘框体"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "旋转转盘窗体.frx":0000
   StartUpPosition =   1  '所有者中心
End
Attribute VB_Name = "旋转转盘框体"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim spinNum As Long
Dim spinWheel As String
Dim i As Long
Dim j As Long

Private Sub UserForm_Initialize()
    Dim arr() As Variant
    arr = Array("白银转盘", "黄金转盘", "钻石转盘")
    ComboBox1.List = Application.Transpose(arr)
End Sub

Private Sub ComboBox1_Change()
    spinWheel = ComboBox1.Value
End Sub

Private Sub TextBox1_Change()
    If IsNumeric(TextBox1.Value) Then
        spinNum = CLng(TextBox1.Value)
    Else
        MsgBox "请输入有效的数字！", vbExclamation
        TextBox1.Value = "" ' 清空文本框
        spinNum = 0 ' 重置 spinNum
    End If
End Sub

Private Sub CommandButton1_Click()
    Dim zuansRewardArr(1 To 12) As Double
    Dim zuansProbabilityArr(1 To 12) As Double
    Dim baiyRewardArr(1 To 12) As Double
    Dim baiyProbabilityArr(1 To 12) As Double
    Dim huangjRewardArr(1 To 12) As Double
    Dim huangjProbabilityArr(1 To 12) As Double
    Dim result As Double
    Dim foundResult As Boolean
    Dim tRow As Long

    ' 初始化随机数生成器
    Randomize

    ' 读取奖品&概率数据
    For i = 1 To 12
        baiyRewardArr(i) = ThisWorkbook.Sheets("转盘").Range("E3").Offset(0, i - 1).Value
        baiyProbabilityArr(i) = ThisWorkbook.Sheets("转盘").Range("E4").Offset(0, i - 1).Value
        huangjRewardArr(i) = ThisWorkbook.Sheets("转盘").Range("E9").Offset(0, i - 1).Value
        huangjProbabilityArr(i) = ThisWorkbook.Sheets("转盘").Range("E10").Offset(0, i - 1).Value
        zuansRewardArr(i) = ThisWorkbook.Sheets("转盘").Range("E15").Offset(0, i - 1).Value
        zuansProbabilityArr(i) = ThisWorkbook.Sheets("转盘").Range("E16").Offset(0, i - 1).Value
    Next i

    For i = 1 To spinNum
        Dim randNum As Double
        Dim cumProb As Double

        randNum = Rnd ' 生成随机数
        cumProb = 0
        foundResult = False ' 重置标志位

        For j = 1 To 12
            If spinWheel = "白银转盘" Then
                cumProb = cumProb + baiyProbabilityArr(j)

                If randNum <= cumProb Then
                    result = baiyRewardArr(j)
                    foundResult = True
                    Exit For
                End If
            ElseIf spinWheel = "黄金转盘" Then
                cumProb = cumProb + huangjProbabilityArr(j)
                 If randNum <= cumProb Then
                    result = huangjRewardArr(j)
                    foundResult = True
                    Exit For
                End If

            ElseIf spinWheel = "钻石转盘" Then
                cumProb = cumProb + zuansProbabilityArr(j)
                 If randNum <= cumProb Then
                    result = zuansRewardArr(j)
                    foundResult = True
                    Exit For
                End If
            End If
        Next j

        If Not foundResult Then
            result = 0 ' 或者其他默认值，或者报错提示
            MsgBox "未找到匹配的奖品，请检查概率设置！", vbCritical
        End If

        ' 找到第一个空白的 T列位置
        tRow = 2
        Do While Not IsEmpty(ThisWorkbook.Sheets("转盘").Range("T" & tRow))
            tRow = tRow + 1
            If tRow > 1000 Then Exit Do ' 使用 Exit Do 而不是 Exit Sub
        Loop

        ' 在找到的空白行写入
        ThisWorkbook.Sheets("转盘").Range("T" & tRow).Value = "您这次" & CStr(spinWheel) & "获得 " & CStr(result) & " 个小票"
    Next i

    旋转转盘框体.Hide
End Sub

