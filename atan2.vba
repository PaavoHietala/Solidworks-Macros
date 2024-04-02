Function ArcTan2(y As Double, x As Double) As Double

    Const pi As Double = 3.14159265

    If x > 0 Then
        ArcTan2 = Atn(y / x)
    ElseIf x < 0 And y >= 0 Then
        ArcTan2 = Atn(y / x) + pi
    ElseIf x < 0 And y < 0 Then
        ArcTan2 = Atn(y / x) - pi
    ElseIf x = 0 And y > 0 Then
        ArcTan2 = pi / 2
    ElseIf x = 0 And y < 0 Then
        ArcTan2 = -pi / 2
    ElseIf x = 0 And y = 0 Then
        ArcTan2 = 0 ' undefined, return 0
    End If

End Function