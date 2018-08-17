Sub SwapPosition()

  Dim shapeRef As shape
  Dim shape1 As shape
  Dim shape2 As shape
  Dim index As Double: index = 0

  Dim shape1Top As Single: shape1Top = 0
  Dim shape1Left As Single: shape1Left = 0

  Dim shape2Top As Single: shape2Top = 0
  Dim shape2Left As Single: shape2Left = 0

  On Error GoTo ERR_HNDL

  For Each shapeRef In Selection.ShapeRange
    index = index + 1
    If index = 1 Then
        Set shape1 = shapeRef
        shape1Top = shape1.Top
        shape1Left = shape1.Left

    ElseIf index = 2 Then
        Set shape2 = shapeRef
        shape2Top = shape2.Top
        shape2Left = shape2.Left
    End If
  Next shapeRef
  shape1.Top = shape2Top
  shape1.Left = shape2Left

  shape2.Top = shape1Top
  shape2.Left = shape1Left

  Exit Sub

ERR_HNDL:
    MsgBox "図形が選択されてないかも"

End Sub
