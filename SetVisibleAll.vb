Sub SetVisibleAll()
  Dim shp As shape
  On Error GoTo ERR_HNDL

  For Each shp In ActiveSheet.Shapes
    shp.Visible = True
  Next shp
  Exit Sub

ERR_HNDL:
    MsgBox "図形が存在しません。"
End Sub
