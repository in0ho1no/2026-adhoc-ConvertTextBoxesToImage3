'═══════════════════════════════════════════════════════
' メイン実行：アクティブブックの全シートを対象に一括処理
'═══════════════════════════════════════════════════════
Sub RunAll()
    Dim wb          As Workbook
    Dim ws          As Worksheet
    Dim totalSheets As Integer
    Dim currentIdx  As Integer

    Set wb = ActiveWorkbook  ' マクロ保存ブックではなくアクティブブックを対象にする

    Application.ScreenUpdating = False
    Application.EnableEvents   = False
    Application.Calculation    = xlCalculationManual

    On Error GoTo ErrHandler

    totalSheets = wb.Worksheets.Count
    currentIdx  = 0

    For Each ws In wb.Worksheets
        currentIdx = currentIdx + 1

        ' 進捗をステータスバーに表示
        Application.StatusBar = "処理中... [" & currentIdx & " / " & totalSheets & "] " & ws.Name

        ws.Activate
        Call ConvertLinkedPicturesToImages(ws)
        Call ConvertShapesToImages(ws)
    Next ws

    Application.StatusBar = False  ' ステータスバーを元に戻す

    Application.ScreenUpdating = True
    Application.EnableEvents   = True
    Application.Calculation    = xlCalculationAutomatic

    MsgBox "✅ 全シートの処理が完了しました。（" & totalSheets & " シート）", vbInformation, "処理完了"
    Exit Sub

ErrHandler:
    Application.StatusBar     = False
    Application.ScreenUpdating = True
    Application.EnableEvents   = True
    Application.Calculation    = xlCalculationAutomatic
    MsgBox "❌ RunAll エラー（" & Err.Number & "）：" & Err.Description, vbCritical, "エラー"
End Sub

'═══════════════════════════════════════════════════════
' カメラ画像（Type=13）→ 通常画像に変換
'═══════════════════════════════════════════════════════
Sub ConvertLinkedPicturesToImages(ws As Worksheet)

    Dim shp     As Shape
    Dim stepMsg As String
    Dim L As Double, T As Double, W As Double, H As Double
    Dim lpNames() As String
    Dim lpCount   As Integer
    Dim j         As Integer
    Dim newPic    As Shape

    On Error GoTo ErrHandler

    stepMsg = "[" & ws.Name & "] カメラ画像の名前リスト取得中"
    lpCount = 0

    For Each shp In ws.Shapes
        If shp.Type = 13 Then
            lpCount = lpCount + 1
            ReDim Preserve lpNames(lpCount - 1)
            lpNames(lpCount - 1) = shp.Name
        End If
    Next shp

    If lpCount = 0 Then Exit Sub

    For j = 0 To lpCount - 1
        stepMsg = "[" & ws.Name & "] カメラ画像取得中 [" & lpNames(j) & "]"

        On Error Resume Next
        Set shp = ws.Shapes(lpNames(j))
        On Error GoTo ErrHandler
        If shp Is Nothing Then GoTo NextLP

        L = shp.Left
        T = shp.Top
        W = shp.Width
        H = shp.Height

        stepMsg = "[" & ws.Name & "] CopyPicture中 [" & lpNames(j) & "]"
        shp.CopyPicture Appearance:=xlScreen, Format:=xlPicture

        ' 先に貼り付けてから削除する（削除するとクリップボードが消えるため）
        stepMsg = "[" & ws.Name & "] 貼り付け中 [" & lpNames(j) & "]"
        ws.Paste
        Set newPic = ws.Shapes(ws.Shapes.Count)

        stepMsg = "[" & ws.Name & "] 削除中 [" & lpNames(j) & "]"
        shp.Delete
        Set shp = Nothing

        With newPic
            .Left   = L
            .Top    = T
            .Width  = W
            .Height = H
            .Name   = "PIC_" & lpNames(j)
        End With

NextLP:
    Next j
    Exit Sub

ErrHandler:
    MsgBox "❌ エラー（" & Err.Number & "）：" & Err.Description & vbCrLf & _
           "発生箇所：" & stepMsg, vbCritical, "エラー"
    Err.Clear
End Sub

'═══════════════════════════════════════════════════════
' テキストボックス整形 → 長方形置換 → グループ化 → 画像化
'═══════════════════════════════════════════════════════
Sub ConvertShapesToImages(ws As Worksheet)

    Dim shp         As Shape
    Dim dict        As Object
    Dim col         As Collection
    Dim k           As Variant
    Dim varArr()    As Variant
    Dim i           As Integer
    Dim targetShape As Shape
    Dim stepMsg     As String
    Dim L As Double, T As Double, W As Double, H As Double
    Dim tbNames()   As String
    Dim tbCount     As Integer
    Dim j           As Integer
    Dim shpName     As String
    Dim newRect     As Shape
    Dim cellAddr    As String
    Dim newPic      As Shape

    Set dict = CreateObject("Scripting.Dictionary")

    On Error GoTo ErrHandler

    '═══════════════════════════
    ' STEP 1-3: テキストボックス → 長方形置換
    '═══════════════════════════
    stepMsg = "[" & ws.Name & "] テキストボックス名リスト取得中"
    tbCount = 0

    For Each shp In ws.Shapes
        If shp.Type = msoTextBox Then
            tbCount = tbCount + 1
            ReDim Preserve tbNames(tbCount - 1)
            tbNames(tbCount - 1) = shp.Name
        End If
    Next shp

    For j = 0 To tbCount - 1
        shpName = tbNames(j)
        stepMsg = "[" & ws.Name & "] STEP1: 引き伸ばし中 [" & shpName & "]"

        On Error Resume Next
        Set shp = ws.Shapes(shpName)
        On Error GoTo ErrHandler
        If shp Is Nothing Then GoTo NextTB

        shp.Width  = shp.Width  * 3
        shp.Height = shp.Height * 3

        stepMsg = "[" & ws.Name & "] STEP2: AutoSize中 [" & shpName & "]"
        shp.TextFrame2.AutoSize = msoAutoSizeShapeToFitText

        stepMsg = "[" & ws.Name & "] STEP3: 長方形作成中 [" & shpName & "]"
        Set newRect = ws.Shapes.AddShape( _
            msoShapeRectangle, _
            shp.Left, shp.Top, shp.Width, shp.Height)

        With newRect
            .TextFrame2.TextRange.Text = shp.TextFrame2.TextRange.Text

            ' フォントカラーを黒に指定（全文字に適用）
            With .TextFrame2.TextRange.Font
                .Fill.ForeColor.RGB = RGB(0, 0, 0)
                .Fill.Visible       = msoTrue
                .Fill.Solid
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

        stepMsg = "[" & ws.Name & "] STEP3: 元テキストボックス削除 [" & shpName & "]"
        shp.Delete
        Set shp = Nothing

NextTB:
    Next j

    '═══════════════════════════
    ' STEP 4: TopLeftCell ごとに集約
    '═══════════════════════════
    stepMsg = "[" & ws.Name & "] STEP4: 図形の集約中"
    For Each shp In ws.Shapes
        If shp.Type <> msoPicture And shp.Type <> msoLinkedPicture Then
            cellAddr = shp.TopLeftCell.Address
            If Not dict.Exists(cellAddr) Then
                dict.Add cellAddr, New Collection
            End If
            dict(cellAddr).Add shp.Name
        End If
    Next shp

    '═══════════════════════════
    ' STEP 5: グループ化 → 画像化 → 置き換え
    '═══════════════════════════
    For Each k In dict.Keys
        Set col = dict(k)
        If col.Count = 0 Then GoTo NextKey

        ReDim varArr(col.Count - 1)
        For i = 1 To col.Count
            varArr(i - 1) = col(i)
        Next i

        stepMsg = "[" & ws.Name & "] STEP5: グループ化中 [セル " & k & " / 図形数:" & col.Count & "]"
        If col.Count = 1 Then
            Set targetShape = ws.Shapes(varArr(0))
        Else
            ws.Shapes(varArr(0)).Select Replace:=True
            For i = 1 To UBound(varArr)
                ws.Shapes(varArr(i)).Select Replace:=False
            Next i
            Set targetShape = Selection.ShapeRange.Group
        End If

        L = targetShape.Left
        T = targetShape.Top
        W = targetShape.Width
        H = targetShape.Height

        stepMsg = "[" & ws.Name & "] STEP5: CopyPicture中 [セル " & k & "]"
        targetShape.CopyPicture Appearance:=xlScreen, Format:=xlPicture

        ' 先に貼り付けてから削除する（削除するとクリップボードが消えるため）
        stepMsg = "[" & ws.Name & "] STEP5: Paste中 [セル " & k & "]"
        ws.Paste
        Set newPic = ws.Shapes(ws.Shapes.Count)

        stepMsg = "[" & ws.Name & "] STEP5: 元図形削除中 [セル " & k & "]"
        targetShape.Delete

        With newPic
            .Left   = L
            .Top    = T
            .Width  = W
            .Height = H
            .Name   = "IMG_" & Replace(k, "$", "")
        End With

NextKey:
    Next k
    Exit Sub

ErrHandler:
    MsgBox "❌ エラー（" & Err.Number & "）：" & Err.Description & vbCrLf & _
           "発生箇所：" & stepMsg, vbCritical, "エラー"
    Err.Clear
End Sub
