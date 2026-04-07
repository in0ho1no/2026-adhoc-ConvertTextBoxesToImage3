'═══════════════════════════════════════════════════════
' フォルダ内の全 .xlsx / .xlsm を対象に一括処理
'═══════════════════════════════════════════════════════
Sub RunAllInFolder()
    Dim fd         As FileDialog
    Dim folderPath As String
    Dim fileName   As String
    Dim wb         As Workbook
    Dim fileCount  As Integer
    Dim errFiles   As String

    ' フォルダ選択はInteractive無効化の前に行う
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    fd.Title = "処理対象フォルダを選択してください"
    If fd.Show <> -1 Then
        MsgBox "フォルダが選択されませんでした。処理を中断します。", vbExclamation, "キャンセル"
        Exit Sub
    End If
    folderPath = fd.SelectedItems(1)
    If Right(folderPath, 1) <> "\" Then folderPath = folderPath & "\"

    ' ファイルループ中の誤操作を防ぐためキー・マウス入力を無効化
    Application.Interactive    = False
    Application.ScreenUpdating = False

    On Error GoTo ErrHandler

    fileCount = 0
    errFiles  = ""

    Dim extensions(1) As String
    extensions(0) = "*.xlsx"
    extensions(1) = "*.xlsm"

    Dim e As Integer
    For e = 0 To 1
        fileName = Dir(folderPath & extensions(e))
        Do While fileName <> ""
            fileCount = fileCount + 1
            Application.StatusBar = "ファイルを開いています... " & fileName

            On Error Resume Next
            Set wb = Workbooks.Open(folderPath & fileName)
            On Error GoTo ErrHandler

            If Not wb Is Nothing Then
                Call RunAll(wb)
                wb.Save
                wb.Close SaveChanges:=False
                Set wb = Nothing
                DoEvents                          ' ファイルClose後にメモリ解放を促す
            Else
                errFiles = errFiles & fileName & vbCrLf
            End If

            fileName = Dir()
        Loop
    Next e

    Application.Interactive    = True   ' ← 必ず元に戻す
    Application.ScreenUpdating = True
    Application.StatusBar      = False

    If errFiles <> "" Then
        MsgBox "⚠️ 以下のファイルでエラーが発生しました：" & vbCrLf & errFiles, _
               vbExclamation, "一部エラー"
    Else
        MsgBox "✅ フォルダ内の全ファイルの処理が完了しました。（" & fileCount & " ファイル）", _
               vbInformation, "処理完了"
    End If
    Exit Sub

ErrHandler:
    ' Interactive=False のまま終了するとExcelが操作不能になるため必ず戻す
    Application.Interactive    = True
    Application.ScreenUpdating = True
    Application.StatusBar      = False
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    MsgBox "❌ RunAllInFolder エラー（" & Err.Number & "）：" & Err.Description, _
           vbCritical, "エラー"
End Sub

'═══════════════════════════════════════════════════════
' 指定ブックの全シートを処理（RunAllInFolder から呼び出す内部処理）
'═══════════════════════════════════════════════════════
Sub RunAll(wb As Workbook)
    Dim ws          As Worksheet
    Dim totalSheets As Integer
    Dim currentIdx  As Integer

    If wb Is Nothing Then Exit Sub

    Application.ScreenUpdating  = False
    Application.EnableEvents    = False
    Application.Calculation     = xlCalculationManual
    Application.EnableCancelKey = xlErrorHandler    ' Escキーをエラーハンドラーへルーティング

    On Error GoTo ErrHandler

    totalSheets = wb.Worksheets.Count
    currentIdx  = 0

    For Each ws In wb.Worksheets
        currentIdx = currentIdx + 1
        Application.StatusBar = "処理中... [" & currentIdx & " / " & totalSheets & "] " & ws.Name

        ' 処理前にA1を選択して図形の選択状態を解除        【2点目：追加】
        ws.Activate
        ws.Cells(1, 1).Select

        Call ConvertLinkedPicturesToImages(ws)
        Call ConvertShapesToImages(ws)

        ' 処理後にA1を選択して最終表示を整える             【3点目：追加】
        ws.Cells(1, 1).Select
    Next ws

    Application.StatusBar       = False
    Application.ScreenUpdating  = True
    Application.EnableEvents    = True
    Application.Calculation     = xlCalculationAutomatic
    Application.EnableCancelKey = xlInterrupt       ' Escキーの挙動をデフォルトに戻す

    Exit Sub

ErrHandler:
    Application.StatusBar       = False
    Application.ScreenUpdating  = True
    Application.EnableEvents    = True
    Application.Calculation     = xlCalculationAutomatic
    Application.EnableCancelKey = xlInterrupt       ' Escキーの挙動をデフォルトに戻す
    MsgBox "❌ RunAll エラー（" & Err.Number & "）：" & Err.Description, vbCritical, "エラー"
End Sub

'═══════════════════════════════════════════════════════
' カメラ画像（Type=13）→「図」に変換
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

        ws.Activate

        ' ── Copy → Paste（最大3回リトライ） ──────────────────────────
        Dim retryLP  As Integer
        Dim pastedLP As Boolean
        retryLP  = 0
        pastedLP = False

        Do While retryLP < 3 And Not pastedLP
            stepMsg = "[" & ws.Name & "] Copy中 [" & lpNames(j) & "] (試行" & retryLP + 1 & ")"
            shp.Copy
            DoEvents                            ' クリップボードが確定するまで処理を一度手放す

            stepMsg = "[" & ws.Name & "] 貼り付け中 [" & lpNames(j) & "] (試行" & retryLP + 1 & ")"
            On Error Resume Next
            ws.Paste
            If Err.Number = 0 Then
                pastedLP = True
            Else
                Err.Clear
                retryLP = retryLP + 1
            End If
            On Error GoTo ErrHandler
        Loop

        If Not pastedLP Then
            MsgBox "❌ Paste失敗（3回リトライ後）：" & lpNames(j), vbCritical, "エラー"
            GoTo NextLP
        End If
        ' ── ここまでリトライ処理 ────────────────────────────────────

        Set newPic = ws.Shapes(ws.Shapes.Count)
        Application.CutCopyMode = False         ' クリップボードを即座に解放

        ' Copy後もリンクが保持されたまま貼り付けられた場合（Type=13）は
        ' 削除してCopyPictureでフォールバック
        If newPic.Type = 13 Then
            newPic.Delete
            Set newPic = Nothing

            Dim retryFB  As Integer
            Dim pastedFB As Boolean
            retryFB  = 0
            pastedFB = False

            Do While retryFB < 3 And Not pastedFB
                stepMsg = "[" & ws.Name & "] フォールバック: CopyPicture中 [" & lpNames(j) & "] (試行" & retryFB + 1 & ")"
                shp.CopyPicture Appearance:=xlScreen, Format:=xlPicture
                DoEvents

                stepMsg = "[" & ws.Name & "] フォールバック: 貼り付け中 [" & lpNames(j) & "] (試行" & retryFB + 1 & ")"
                On Error Resume Next
                ws.Paste
                If Err.Number = 0 Then
                    pastedFB = True
                Else
                    Err.Clear
                    retryFB = retryFB + 1
                End If
                On Error GoTo ErrHandler
            Loop

            If Not pastedFB Then
                MsgBox "❌ フォールバックPaste失敗（3回リトライ後）：" & lpNames(j), vbCritical, "エラー"
                GoTo NextLP
            End If

            Set newPic = ws.Shapes(ws.Shapes.Count)
            Application.CutCopyMode = False     ' クリップボードを即座に解放
        End If

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
' テキストボックス整形 → グループ化 → 画像化
' ※長方形への置換を廃止し、書式を保持したままグループ化   【1点目：STEP3を削除】
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
    Dim cellAddr    As String
    Dim newPic      As Shape

    Set dict = CreateObject("Scripting.Dictionary")

    On Error GoTo ErrHandler

    '═══════════════════════════
    ' STEP 1-2: テキストボックスの見切れ対処
    ' ※長方形への置換は行わず、書式をそのまま保持         【1点目：STEP3を削除】
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

        ' 幅を3倍にしてから AutoSize することで、
        ' テキストの折り返し基準幅を広げつつ高さを自動調整する
        shp.Width  = shp.Width * 3

        stepMsg = "[" & ws.Name & "] STEP2: AutoSize中 [" & shpName & "]"
        shp.TextFrame2.AutoSize = msoAutoSizeShapeToFitText

        ' ここで長方形への置換は行わない（書式保持のため）  【1点目】

NextTB:
    Next j

    '═══════════════════════════
    ' STEP 3: TopLeftCell ごとに集約
    '═══════════════════════════
    stepMsg = "[" & ws.Name & "] STEP3: 図形の集約中"
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
    ' STEP 4: グループ化 → 画像化 → 置き換え
    '═══════════════════════════
    For Each k In dict.Keys
        Set col = dict(k)
        If col.Count = 0 Then GoTo NextKey

        ReDim varArr(col.Count - 1)
        For i = 1 To col.Count
            varArr(i - 1) = col(i)
        Next i

        stepMsg = "[" & ws.Name & "] STEP4: グループ化中 [セル " & k & " / 図形数:" & col.Count & "]"
        ws.Activate
        If col.Count = 1 Then
            Set targetShape = ws.Shapes(varArr(0))
        Else
            ' Select を使わず ShapeRange を直接生成してグループ化
            ' → Select/Replace:=False の失敗を根本回避
            Set targetShape = ws.Shapes.Range(varArr).Group
        End If

        L = targetShape.Left
        T = targetShape.Top
        W = targetShape.Width
        H = targetShape.Height

        ws.Activate

        ' ── CopyPicture → Paste（最大3回リトライ） ──────────────────
        Dim retrySH  As Integer
        Dim pastedSH As Boolean
        retrySH  = 0
        pastedSH = False

        Do While retrySH < 3 And Not pastedSH
            stepMsg = "[" & ws.Name & "] STEP4: CopyPicture中 [セル " & k & "] (試行" & retrySH + 1 & ")"
            targetShape.CopyPicture Appearance:=xlScreen, Format:=xlPicture
            DoEvents                            ' クリップボードが確定するまで処理を一度手放す

            stepMsg = "[" & ws.Name & "] STEP4: Paste中 [セル " & k & "] (試行" & retrySH + 1 & ")"
            On Error Resume Next
            ws.Paste
            If Err.Number = 0 Then
                pastedSH = True
            Else
                Err.Clear
                retrySH = retrySH + 1
            End If
            On Error GoTo ErrHandler
        Loop

        If Not pastedSH Then
            MsgBox "❌ Paste失敗（3回リトライ後）：セル " & k, vbCritical, "エラー"
            GoTo NextKey
        End If
        ' ── ここまでリトライ処理 ────────────────────────────────────

        Set newPic = ws.Shapes(ws.Shapes.Count)
        Application.CutCopyMode = False         ' クリップボードを即座に解放

        stepMsg = "[" & ws.Name & "] STEP4: 元図形削除中 [セル " & k & "]"
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
