Sub DiagnoseShapes()

    Dim ws  As Worksheet
    Dim shp As Shape
    Dim msg As String

    Set ws = ActiveSheet
    msg = "図形名 / Type値 / Type名" & vbCrLf
    msg = msg & String(60, "-") & vbCrLf

    For Each shp In ws.Shapes
        Dim typeName As String
        Select Case shp.Type
            Case 1:  typeName = "msoAutoShape"
            Case 3:  typeName = "msoPicture"
            Case 8:  typeName = "msoLinkedOLEObject"
            Case 11: typeName = "msoLinkedPicture"
            Case 12: typeName = "msoOLEControlObject"
            Case 13: typeName = "msoPicture(13)"
            Case 17: typeName = "msoTextBox"
            Case 19: typeName = "msoLinkedPicture(19)?"
            Case Else: typeName = "その他(" & shp.Type & ")"
        End Select
        msg = msg & shp.Name & " / " & shp.Type & " / " & typeName & vbCrLf
    Next shp

    ' 長い場合はイミディエイトウィンドウに出力
    Debug.Print msg
    MsgBox msg, vbInformation, "図形タイプ診断"

End Sub
