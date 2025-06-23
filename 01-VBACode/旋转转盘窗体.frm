VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ��תת�̿��� 
   Caption         =   "��תת�̿���"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "��תת�̴���.frx":0000
   StartUpPosition =   1  '����������
End
Attribute VB_Name = "��תת�̿���"
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
    arr = Array("����ת��", "�ƽ�ת��", "��ʯת��")
    ComboBox1.List = Application.Transpose(arr)
End Sub

Private Sub ComboBox1_Change()
    spinWheel = ComboBox1.Value
End Sub

Private Sub TextBox1_Change()
    If IsNumeric(TextBox1.Value) Then
        spinNum = CLng(TextBox1.Value)
    Else
        MsgBox "��������Ч�����֣�", vbExclamation
        TextBox1.Value = "" ' ����ı���
        spinNum = 0 ' ���� spinNum
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

    ' ��ʼ�������������
    Randomize

    ' ��ȡ��Ʒ&��������
    For i = 1 To 12
        baiyRewardArr(i) = ThisWorkbook.Sheets("ת��").Range("E3").Offset(0, i - 1).Value
        baiyProbabilityArr(i) = ThisWorkbook.Sheets("ת��").Range("E4").Offset(0, i - 1).Value
        huangjRewardArr(i) = ThisWorkbook.Sheets("ת��").Range("E9").Offset(0, i - 1).Value
        huangjProbabilityArr(i) = ThisWorkbook.Sheets("ת��").Range("E10").Offset(0, i - 1).Value
        zuansRewardArr(i) = ThisWorkbook.Sheets("ת��").Range("E15").Offset(0, i - 1).Value
        zuansProbabilityArr(i) = ThisWorkbook.Sheets("ת��").Range("E16").Offset(0, i - 1).Value
    Next i

    For i = 1 To spinNum
        Dim randNum As Double
        Dim cumProb As Double

        randNum = Rnd ' ���������
        cumProb = 0
        foundResult = False ' ���ñ�־λ

        For j = 1 To 12
            If spinWheel = "����ת��" Then
                cumProb = cumProb + baiyProbabilityArr(j)

                If randNum <= cumProb Then
                    result = baiyRewardArr(j)
                    foundResult = True
                    Exit For
                End If
            ElseIf spinWheel = "�ƽ�ת��" Then
                cumProb = cumProb + huangjProbabilityArr(j)
                 If randNum <= cumProb Then
                    result = huangjRewardArr(j)
                    foundResult = True
                    Exit For
                End If

            ElseIf spinWheel = "��ʯת��" Then
                cumProb = cumProb + zuansProbabilityArr(j)
                 If randNum <= cumProb Then
                    result = zuansRewardArr(j)
                    foundResult = True
                    Exit For
                End If
            End If
        Next j

        If Not foundResult Then
            result = 0 ' ��������Ĭ��ֵ�����߱�����ʾ
            MsgBox "δ�ҵ�ƥ��Ľ�Ʒ������������ã�", vbCritical
        End If

        ' �ҵ���һ���հ׵� T��λ��
        tRow = 2
        Do While Not IsEmpty(ThisWorkbook.Sheets("ת��").Range("T" & tRow))
            tRow = tRow + 1
            If tRow > 1000 Then Exit Do ' ʹ�� Exit Do ������ Exit Sub
        Loop

        ' ���ҵ��Ŀհ���д��
        ThisWorkbook.Sheets("ת��").Range("T" & tRow).Value = "�����" & CStr(spinWheel) & "��� " & CStr(result) & " ��СƱ"
    Next i

    ��תת�̿���.Hide
End Sub

