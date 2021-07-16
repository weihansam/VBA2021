Attribute VB_Name = "Module1"
Sub bmi()

    Dim h!
    Dim w!
    Dim bmi!
    Dim age%
    h = Range("B1").Value
    w = Range("B2").Value
    age = Range("D2").Value
    bmi = w / ((h / 100) ^ 2)
    Range("B3").Value = bmi

End Sub

