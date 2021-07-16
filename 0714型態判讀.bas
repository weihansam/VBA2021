Attribute VB_Name = "Module1"
Option Explicit

Sub typValidate()


Range("A2").Value = TypeName(Range("A1").Value)

'ゅ綼オ 计綼 booleanい丁

End Sub

Sub typValidate2()


Range("A2").Value = TypeName(Range("A1").Value)

'If IsNumeric(Range("A1").Value) Then
'MsgBox ("琌计篈")
'Else
'MsgBox ("ぃ琌计篈")
'End If
'ゅら戳耞程非

Range("A3").Value = VarType(Range("A1").Value)


End Sub


