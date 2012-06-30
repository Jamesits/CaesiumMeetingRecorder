Attribute VB_Name = "ModTimer"
Option Explicit

Public Function FormatTime(ByVal Sec) As String  '秒数格式化为时：分：秒的字符串
Dim g, h, I As String
g = Int(Sec / 3600)
h = Int((Sec - 3600 * Val(g)) / 60)
I = Int(Sec - 3600 * Val(g) - 60 * Val(h))
If Val(g) < 10 Then g = "0" & g
If Val(h) < 10 Then h = "0" & h
If Val(I) < 10 Then I = "0" & I
FormatTime = g & ":" & h & ":" & I
End Function

