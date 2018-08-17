Sub HideShapes()
  Dim shape As shape
  On Error GoTo ERR_HNDL

  For Each shape In Selection.ShapeRange
    If shape.Visible = True Then
        shape.Visible = False
    End If
  Next shape
  Exit Sub

ERR_HNDL:
    MsgBox "図形が選択されていません。"
End Sub
