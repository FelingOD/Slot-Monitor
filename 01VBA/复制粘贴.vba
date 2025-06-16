Sub 转移数据()

' 转移数据 宏

    dim toCopyPos as long
    dim tpWritePos as long
    Range("D17").Select
    Selection.Copy
    Range("F17").Select
    ActiveSheet.Paste
End Sub
