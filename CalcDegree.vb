Sub CalcDegree()

  Dim shape As shape
  On Error GoTo ERR_HNDL
  Dim angle As Double
  For Each shape In Selection.ShapeRange
    angle = GetAngle(Atn(shape.Height / shape.Width))
    MsgBox "角度:" & Application.RoundDown(angle, 2) & vbCrLf _
                                & "距離:" & Application.RoundDown(shape.Height/Sin(Atn(shape.Height / shape.Width)), 2) & vbCrLf _
                                & "高さ:" & Application.RoundDown(shape.Height, 2) & vbCrLf _
                                & "幅:" & Application.RoundDown(shape.Width, 2)
  Next shape
  Exit Sub

ERR_HNDL:
    MsgBox "図形が選択されてない" & vbCrLf _
            & "or 高さと幅どっちかが0です"
End Sub

Function PI() As Double
    PI = Atn(1) * 4
End Function

Function GetAngle(ByVal radian As Double) As Double
    GetAngle = radian / (PI / 180)
End Function
