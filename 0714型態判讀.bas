Attribute VB_Name = "Module1"
Option Explicit

Sub typValidate()


Range("A2").Value = TypeName(Range("A1").Value)

'��r�a�� �Ʀr�a�k boolean�b����

End Sub

Sub typValidate2()


Range("A2").Value = TypeName(Range("A1").Value)

'If IsNumeric(Range("A1").Value) Then
'MsgBox ("�O�ƭȫ��A")
'Else
'MsgBox ("���O�ƭȫ��A")
'End If
'��r�B����P�_�̷�

Range("A3").Value = VarType(Range("A1").Value)


End Sub


