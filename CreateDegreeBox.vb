Sub CreateDegreeBox()

  Dim shape As shape
  On Error GoTo ERR_HNDL
  Dim angle As Double
  For Each shape In Selection.ShapeRange
    angle = GetAngle(Atn(shape.Height / shape.Width))
    ActiveSheet.Shapes.AddTextbox( _
            Orientation:=msoTextOrientationHorizontal, _
            Left:=shape.Left + 25, _
            Top:=shape.Top - 25, _
            Width:=75, _
            Height:=75).Select
    Selection.Characters.Text = "角度:" & Application.RoundDown(angle, 2) & vbCrLf _
                                & "距離:" & Application.RoundDown(shape.Height/Sin(Atn(shape.Height / shape.Width)), 2) & vbCrLf _
                                & "高さ:" & Application.RoundDown(shape.Height, 2) & vbCrLf _
                                & "幅:" & Application.RoundDown(shape.Width, 2)

  Next shape
  Exit Sub

ERR_HNDL:
    MsgBox "図形が選択されていません。"
End Sub
