Attribute VB_Name = "Module1"
Sub notEQLdeom()

    If (reange("A1").Value = Range("B1").Value) Then
    
    MsgBox ("�۵�")
    
    End If
    
    
    If (reange("A1").Value <> Range("B1").Value) Then

    MsgBox ("���۵�")
    
    End If
    
    
End Sub

Sub combine()
    '�k1
    Dim num1, num2 As Integer
    num1 = 6
    num2 = 8
    Sum = num1 + num2
    MsgBox (Sum)
    
    '�k2
    num1 = 56: num2 = 8
    
    Sum = num1 + num2
    MsgBox (Sum)
    

End Sub
