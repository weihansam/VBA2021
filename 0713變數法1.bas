Attribute VB_Name = "Module1"
Sub notEQLdeom()

    If (reange("A1").Value = Range("B1").Value) Then
    
    MsgBox ("相等")
    
    End If
    
    
    If (reange("A1").Value <> Range("B1").Value) Then

    MsgBox ("不相等")
    
    End If
    
    
End Sub

Sub combine()
    '法1
    Dim num1, num2 As Integer
    num1 = 6
    num2 = 8
    Sum = num1 + num2
    MsgBox (Sum)
    
    '法2
    num1 = 56: num2 = 8
    
    Sum = num1 + num2
    MsgBox (Sum)
    

End Sub
