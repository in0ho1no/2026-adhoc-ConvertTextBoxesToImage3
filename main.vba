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

    Set ws   = ActiveSheet
    Set dict = CreateObject("Scripting.Dictionary")

    Application.ScreenUpdating = False
    Application.EnableEvents   = False
    Application.Calculation    = xlCalculationManual

    On Error GoTo ErrHandler

    '═══════════════════════════════════════════════════════
    ' STEP 1 & 2 & 3: テキストボックスを整形 → 長方形に置換
    '═══════════════════════════════════════════════════════

    ' ループ中に図形を追加・削除するため、先に名前リストを取得する
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

        ' 図形が存在するか確認（処理中に消えた場合のガード）
        Dim shpName As String
        shpName = tbNames(j)
        On Error Resume Next
        Set shp = ws.Shapes(shpName)
        On Error GoTo ErrHandler
        If shp Is Nothing Then GoTo NextTB

        ' ── STEP 1: 縦横3倍に引き伸ばす ───────────────────────
        With shp
            .Width  = .Width  * 3
            .Height = .Height * 3
        End With

        ' ── STEP 2: テキストにフィットさせる ───────────────────
        With shp.TextFrame2
            .AutoSize = msoAutoSizeShapeToFitText
        End With

        ' ── STEP 3: 長方形を新規作成し属性を引き継ぐ ───────────
        Dim newRect As Shape
        Set newRect = ws.Shapes.AddShape( _
            msoShapeRectangle, _
            shp.Left, shp.Top, shp.Width, shp.Height)

        With newRect
            ' テキスト内容
            .TextFrame2.TextRange.Text = shp.TextFrame2.TextRange.Text

            ' フォント設定（Meiryo UI・10・黒）
            With .TextFrame2.TextRange.Font
                .Name = "Meiryo UI"
                .Size = 10
                .Fill.ForeColor.RGB = RGB(0, 0, 0)
            End With

            ' テキスト余白を元と揃える
            With .TextFrame2
                .MarginLeft   = shp.TextFrame2.MarginLeft
                .MarginRight  = shp.TextFrame2.MarginRight
                .MarginTop    = shp.TextFrame2.MarginTop
                .MarginBottom = shp.TextFrame2.MarginBottom
                .WordWrap     = shp.TextFrame2.WordWrap
                ' サイズは固定（AutoSizeは不要）
                .AutoSize     = msoAutoSizeNone
            End With

            ' 背景：透明
            .Fill.Visible = msoFalse

            ' 枠線：非表示（元テキストボックスと同様）
            .Line.Visible = msoFalse

            ' 名前を引き継ぐ（後のグループ化で使用）
            .Name = "RECT_" & shpName
        End With

        ' 元テキストボックスを削除
        shp.Delete
        Set shp = Nothing

NextTB:
    Next j

    '═══════════════════════════════════════════════════════
    ' STEP 4: 全図形を TopLeftCell ごとに集約
    ' ※ すでに画像になっているもの（Picture）はスキップ
    '═══════════════════════════════════════════════════════
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

        ' 図形名を配列化
        ReDim varArr(col.Count - 1)
        For i = 1 To col.Count
            varArr(i - 1) = col(i)
        Next i

        ' 1つだけならそのまま・2つ以上はグループ化
        If col.Count = 1 Then
            Set targetShape = ws.Shapes(varArr(0))
        Else
            Set targetShape = ws.Shapes.Range(varArr).Group
        End If

        ' 位置・サイズを記録
        L = targetShape.Left
        T = targetShape.Top
        W = targetShape.Width
        H = targetShape.Height

        ' 画像としてコピー
        targetShape.CopyPicture Appearance:=xlScreen, Format:=xlPicture

        ' 元図形を削除
        targetShape.Delete

        ' シートに貼り付けて位置・サイズを復元
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

    MsgBox "✅ 完了：テキストボックス整形 → 長方形置換 → 画像化 が終わりました。", _
           vbInformation, "処理完了"

Cleanup:
    Application.ScreenUpdating = True
    Application.EnableEvents   = True
    Application.Calculation    = xlCalculationAutomatic
    Exit Sub

ErrHandler:
    MsgBox "❌ エラー（" & Err.Number & "）：" & Err.Description, _
           vbCritical, "エラー"
    Resume Cleanup

End Sub
