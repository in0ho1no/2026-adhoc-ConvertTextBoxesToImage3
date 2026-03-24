Sub ConvertShapesToImages()

    Dim ws          As Worksheet
    Dim shp         As Shape
    Dim dict        As Object
    Dim col         As Collection
    Dim k           As Variant
    Dim varArr()    As Variant
    Dim i           As Integer
    Dim targetShape As Shape
    Dim L As Double, T As Double, W As Double, H As Double
    Dim stepMsg     As String   ' ← どこで落ちたか追跡用

    Set ws   = ActiveSheet
    Set dict = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False
    Application.EnableEvents   = False
    Application.Calculation    = xlCalculationManual

    On Error GoTo ErrHandler

    '═══════════════════════════════════════════════════════
    ' STEP 1 & 2 & 3: テキストボックス整形 → 長方形置換
    '═══════════════════════════════════════════════════════
    stepMsg = "STEP1-3: テキストボックス名リスト取得"
    Dim tbNames() As String
    Dim tbCount   As Integer
    tbCount = 0

    For Each shp In ws.Shapes
        If shp.Type = msoTextBox Then
            tbCount = tbCount + 1
            ReDim Preserve tbNames(tbCount - 1)
            tbNames(tbCount - 1) = shp.Name
        End If
    Next shp

    Dim j As Integer
    For j = 0 To tbCount - 1
        Dim shpName As String
        shpName = tbNames(j)
        stepMsg = "STEP1: 引き伸ばし中 [" & shpName & "]"

        On Error Resume Next
        Set shp = ws.Shapes(shpName)
        On Error GoTo ErrHandler
        If shp Is Nothing Then GoTo NextTB

        ' STEP1: 3倍引き伸ばし
        shp.Width  = shp.Width  * 3
        shp.Height = shp.Height * 3

        ' STEP2: AutoSize
        stepMsg = "STEP2: AutoSize中 [" & shpName & "]"
        shp.TextFrame2.AutoSize = msoAutoSizeShapeToFitText

        ' STEP3: 長方形を新規作成
        stepMsg = "STEP3: 長方形作成中 [" & shpName & "]"
        Dim newRect As Shape
        Set newRect = ws.Shapes.AddShape( _
            msoShapeRectangle, _
            shp.Left, shp.Top, shp.Width, shp.Height)

        With newRect
            .TextFrame2.TextRange.Text = shp.TextFrame2.TextRange.Text
            With .TextFrame2.TextRange.Font
                .Name = "Meiryo UI"
                .Size = 10
                .Fill.ForeColor.RGB = RGB(0, 0, 0)
            End With
            With .TextFrame2
                .MarginLeft   = shp.TextFrame2.MarginLeft
                .MarginRight  = shp.TextFrame2.MarginRight
                .MarginTop    = shp.TextFrame2.MarginTop
                .MarginBottom = shp.TextFrame2.MarginBottom
                .WordWrap     = shp.TextFrame2.WordWrap
                .AutoSize     = msoAutoSizeNone
            End With
            .Fill.Visible = msoFalse
            .Line.Visible = msoFalse
            .Name = "RECT_" & shpName
        End With

        stepMsg = "STEP3: 元テキストボックス削除 [" & shpName & "]"
        shp.Delete
        Set shp = Nothing

NextTB:
    Next j

    '═══════════════════════════════════════════════════════
    ' STEP 4: TopLeftCell ごとに図形を集約
    '═══════════════════════════════════════════════════════
    stepMsg = "STEP4: 図形の集約中"
    For Each shp In ws.Shapes
        If shp.Type <> msoPicture And shp.Type <> msoLinkedPicture Then
            Dim cellAddr As String
            cellAddr = shp.TopLeftCell.Address
            If Not dict.Exists(cellAddr) Then
                dict.Add cellAddr, New Collection
            End If
            dict(cellAddr).Add shp.Name
        End If
    Next shp

    '═══════════════════════════════════════════════════════
    ' STEP 5: グループ化 → 画像化 → 置き換え
    '═══════════════════════════════════════════════════════
    For Each k In dict.Keys
        Set col = dict(k)
        If col.Count = 0 Then GoTo NextKey

        ReDim varArr(col.Count - 1)
        For i = 1 To col.Count
            varArr(i - 1) = col(i)
        Next i

        stepMsg = "STEP5: グループ化中 [セル " & k & " / 図形数:" & col.Count & "]"
        If col.Count = 1 Then
            Set targetShape = ws.Shapes(varArr(0))
        Else
            Set targetShape = ws.Shapes.Range(varArr).Group
        End If

        L = targetShape.Left
        T = targetShape.Top
        W = targetShape.Width
        H = targetShape.Height

        stepMsg = "STEP5: CopyPicture中 [セル " & k & "]"
        targetShape.CopyPicture Appearance:=xlScreen, Format:=xlPicture

        stepMsg = "STEP5: 元図形削除中 [セル " & k & "]"
        targetShape.Delete

        stepMsg = "STEP5: Paste中 [セル " & k & "]"
        ws.Paste
        With ws.Shapes(ws.Shapes.Count)
            .Left   = L
            .Top    = T
            .Width  = W
            .Height = H
            .Name   = "IMG_" & Replace(k, "$", "")
        End With

NextKey:
    Next k

    MsgBox "✅ 完了しました。", vbInformation, "処理完了"

Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents   = True
    Application.Calculation    = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    MsgBox "❌ エラー（" & Err.Number & "）：" & Err.Description & vbCrLf & vbCrLf & _
           "発生箇所：" & stepMsg, vbCritical, "エラー"
    Resume Cleanup

End Sub
