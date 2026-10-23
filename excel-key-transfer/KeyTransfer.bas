Option Explicit

'============================================================
' キー一致転記マクロ
'   転記元と転記先で「キー」が一致する行について、
'   転記元の指定列の値を、転記先の指定列へ書き込む。
'   （XLOOKUP の結果を値で貼り付けるのと同じ結果）
'
' 使い方：TransferByKey を実行し、画面の案内に従ってセルをクリックするだけ。
'         コードを書き換える必要はない。前回の指定は記憶され、次回そのまま使える。
'============================================================

'==================== 詳細オプション（普段は触らなくてよい） ====================
Private Const SKIP_BLANK_SOURCE As Boolean = False ' True: 転記元が空欄なら転記先を上書きしない
Private Const NORMALIZE_WIDTH As Boolean = True    ' True: キーの全角/半角・前後の空白を無視して照合
Private Const MARK_UNMATCHED As Boolean = True     ' True: 転記元に見つからなかった転記先のキーセルを黄色にする
'==============================================================================

Private Const TITLE_TEXT As String = "キー一致転記"
Private Const REG_APP As String = "KeyTransferMacro"
Private Const REG_SECTION As String = "LastSettings"

Public Sub TransferByKey()
    Dim srcWb As Workbook, dstWb As Workbook
    Dim srcOpenedHere As Boolean, dstOpenedHere As Boolean
    Dim srcKey As Range, dstKey As Range   ' キー列の「最初のデータセル」
    Dim srcCols() As Long, dstCols() As Long, nMap As Long
    Dim ans As VbMsgBoxResult, useLast As Boolean

    On Error GoTo Fail

    Set srcWb = PickWorkbook("【1/2】転記元（値を持ってくる側）のファイルを選んでください", srcOpenedHere)
    If srcWb Is Nothing Then Exit Sub
    Set dstWb = PickWorkbook("【2/2】転記先（値を書き込む側）のファイルを選んでください", dstOpenedHere)
    If dstWb Is Nothing Then GoTo Cleanup

    ' 前回の指定が使えるなら、それを使うか聞く
    If LoadSettings(srcWb, dstWb, srcKey, dstKey, srcCols, dstCols, nMap) Then
        ans = MsgBox("前回と同じ指定で転記しますか？" & vbCrLf & vbCrLf & _
                     Describe(srcKey, dstKey, srcCols, dstCols, nMap) & vbCrLf & _
                     "［はい］このまま転記" & vbCrLf & _
                     "［いいえ］画面で指定し直す" & vbCrLf & _
                     "［キャンセル］やめる", _
                     vbYesNoCancel + vbQuestion, TITLE_TEXT)
        If ans = vbCancel Then GoTo Cleanup
        useLast = (ans = vbYes)
    End If

    If Not useLast Then
        If Not AskSettings(srcWb, dstWb, srcKey, dstKey, srcCols, dstCols, nMap) Then GoTo Cleanup
        ans = MsgBox("この内容で転記します。よろしいですか？" & vbCrLf & vbCrLf & _
                     Describe(srcKey, dstKey, srcCols, dstCols, nMap), _
                     vbOKCancel + vbQuestion, TITLE_TEXT)
        If ans <> vbOK Then GoTo Cleanup
    End If

    SaveSettings srcKey, dstKey, srcCols, dstCols, nMap
    DoTransfer srcKey, dstKey, srcCols, dstCols, nMap
    ShowSheet dstKey.Worksheet

Cleanup:
    On Error Resume Next
    ' このマクロが開いた転記元は保存せずに閉じる（転記先と同じファイルなら閉じない）
    If srcOpenedHere And Not srcWb Is Nothing Then
        If dstWb Is Nothing Then
            srcWb.Close SaveChanges:=False
        ElseIf srcWb.FullName <> dstWb.FullName Then
            srcWb.Close SaveChanges:=False
        End If
    End If
    Exit Sub

Fail:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, TITLE_TEXT
    Resume Cleanup
End Sub

'------------------------------------------------------------
' 画面で指定してもらう
'------------------------------------------------------------
Private Function AskSettings(ByVal srcWb As Workbook, ByVal dstWb As Workbook, _
                            ByRef srcKey As Range, ByRef dstKey As Range, _
                            ByRef srcCols() As Long, ByRef dstCols() As Long, ByRef n As Long) As Boolean
    Dim c As Range, d As Range, i As Long, dupe As Boolean

    ' 手順1：転記元のキー
    srcWb.Activate
    Set srcKey = PickCell("【手順1】転記元のキー" & vbCrLf & vbCrLf & _
                          "キー（突き合わせに使う列）の、見出しのすぐ下のセルをクリックして［OK］。" & vbCrLf & _
                          "別のシートなら、先にシート見出しで切り替えてください。", _
                          srcWb, Nothing)
    If srcKey Is Nothing Then Exit Function

    ' 手順2：転記先のキー
    dstWb.Activate
    Set dstKey = PickCell("【手順2】転記先のキー" & vbCrLf & vbCrLf & _
                          "キーの列の、見出しのすぐ下のセルをクリックして［OK］。" & vbCrLf & _
                          "（転記元と別の列でもかまいません）", _
                          dstWb, Nothing)
    If dstKey Is Nothing Then Exit Function

    ' 手順3・4：転記する列の組を、キャンセルされるまで繰り返し聞く
    n = 0
    ReDim srcCols(1 To 1)
    ReDim dstCols(1 To 1)
    Do
        ShowSheet srcKey.Worksheet
        Set c = PickCell("【手順3】転記したい列（転記元）　" & n + 1 & "組目" & vbCrLf & vbCrLf & _
                         "持ってきたい値がある列のセルを、どれか1つクリックして［OK］。" & _
                         IIf(n = 0, "", vbCrLf & "指定済み: " & n & "組。全部指定し終わったら［キャンセル］。"), _
                         srcWb, srcKey.Worksheet)
        If c Is Nothing Then
            If n > 0 Then Exit Do
            If MsgBox("転記する列がまだ1つも指定されていません。やめますか？", vbYesNo + vbQuestion, TITLE_TEXT) = vbYes Then Exit Function
        Else
            ShowSheet dstKey.Worksheet
            Set d = PickCell("【手順4】書き込み先の列（転記先）" & vbCrLf & vbCrLf & _
                             "転記元の " & ColLabel(srcKey.Worksheet, c.Column, srcKey.Row) & " を、" & vbCrLf & _
                             "どの列に書き込みますか？ 列のセルをどれか1つクリックして［OK］。" & vbCrLf & _
                             "（［キャンセル］でこの組は取り消し）", _
                             dstWb, dstKey.Worksheet)
            If Not d Is Nothing Then
                dupe = False
                For i = 1 To n
                    If dstCols(i) = d.Column Then dupe = True
                Next i
                If d.Column = dstKey.Column Then
                    MsgBox "キーの列には書き込めません。別の列を選んでください。", vbExclamation, TITLE_TEXT
                ElseIf dupe Then
                    MsgBox ColLetter(d.Column) & "列はすでに書き込み先に指定済みです。別の列を選んでください。", vbExclamation, TITLE_TEXT
                Else
                    n = n + 1
                    ReDim Preserve srcCols(1 To n)
                    ReDim Preserve dstCols(1 To n)
                    srcCols(n) = c.Column
                    dstCols(n) = d.Column
                End If
            End If
        End If
    Loop
    AskSettings = True
End Function

' セルを1つクリックしてもらう。キャンセルなら Nothing。
' （Application.InputBox は案内文が長すぎると失敗するので、案内文は短く保つこと）
' wb（と ws）以外のセルが選ばれたら、注意して選び直してもらう。
Private Function PickCell(ByVal prompt As String, ByVal wb As Workbook, ByVal ws As Worksheet) As Range
    Dim r As Range
    Do
        Set r = Nothing
        On Error Resume Next
        Set r = Application.InputBox(prompt, TITLE_TEXT, Type:=8)
        On Error GoTo 0
        If r Is Nothing Then Exit Function

        If r.Worksheet.Parent.FullName <> wb.FullName Then
            MsgBox "別のファイルのセルが選ばれています。" & vbCrLf & _
                   "「" & wb.Name & "」のセルを選んでください。", vbExclamation, TITLE_TEXT
            If ws Is Nothing Then wb.Activate Else ShowSheet ws
        ElseIf Not ws Is Nothing Then
            If r.Worksheet.Name <> ws.Name Then
                MsgBox "シート「" & ws.Name & "」のセルを選んでください。" & vbCrLf & _
                       "（キーと同じシートの列だけ指定できます）", vbExclamation, TITLE_TEXT
                ShowSheet ws
            Else
                Set PickCell = r.Cells(1, 1)
                Exit Function
            End If
        Else
            Set PickCell = r.Cells(1, 1)
            Exit Function
        End If
    Loop
End Function

' ファイル選択。既に開いていればそれを使い、開いていなければ開く。キャンセル時は Nothing。
Private Function PickWorkbook(ByVal prompt As String, ByRef openedHere As Boolean) As Workbook
    Dim f As Variant, wb As Workbook
    openedHere = False
    MsgBox prompt, vbInformation, TITLE_TEXT
#If Mac Then
    f = Application.GetOpenFilename(Title:=prompt)
#Else
    f = Application.GetOpenFilename("Excel ファイル (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", , prompt)
#End If
    If VarType(f) = vbBoolean Then Exit Function

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, CStr(f), vbTextCompare) = 0 Then
            Set PickWorkbook = wb
            Exit Function
        End If
    Next wb
    Set PickWorkbook = Workbooks.Open(Filename:=CStr(f))
    openedHere = True
End Function

Private Sub ShowSheet(ByVal ws As Worksheet)
    ws.Parent.Activate
    ws.Activate
End Sub

'------------------------------------------------------------
' 転記本体
'------------------------------------------------------------
Private Sub DoTransfer(ByVal srcKey As Range, ByVal dstKey As Range, _
                       ByRef srcCols() As Long, ByRef dstCols() As Long, ByVal nMap As Long)
    Dim srcWs As Worksheet, dstWs As Worksheet
    Dim srcLast As Long, dstLast As Long
    Dim srcKeys As Variant, dstKeys As Variant, srcData() As Variant
    Dim keyIndex As Collection, isFirst() As Boolean, used() As Boolean
    Dim i As Long, m As Long, r As Long, srcIdx As Long
    Dim k As String, v As Variant
    Dim nMatched As Long, nUnmatched As Long, nBlankKey As Long, nDup As Long, nUnused As Long
    Dim appChanged As Boolean, prevCalc As XlCalculation
    Dim errNum As Long, errDesc As String

    Set srcWs = srcKey.Worksheet
    Set dstWs = dstKey.Worksheet

    On Error GoTo Finish

    ' ---- 転記元を読み込み、キー → 行番号 の索引を作る ----
    srcLast = srcWs.Cells(srcWs.Rows.Count, srcKey.Column).End(xlUp).Row
    If srcLast < srcKey.Row Then
        MsgBox "転記元のキーの列（" & ColLetter(srcKey.Column) & "列）にデータがありません。", vbExclamation, TITLE_TEXT
        Exit Sub
    End If
    srcKeys = ReadCol(srcWs, srcKey.Column, srcKey.Row, srcLast)
    ReDim srcData(1 To nMap)
    For m = 1 To nMap
        srcData(m) = ReadCol(srcWs, srcCols(m), srcKey.Row, srcLast)
    Next m

    Set keyIndex = New Collection
    ReDim isFirst(1 To UBound(srcKeys, 1))
    ReDim used(1 To UBound(srcKeys, 1))
    For i = 1 To UBound(srcKeys, 1)
        k = NormKey(srcKeys(i, 1))
        If Len(k) > 0 Then
            If LookupIndex(keyIndex, k) = 0 Then
                keyIndex.Add i, k
                isFirst(i) = True
            Else
                nDup = nDup + 1   ' 重複キーは先に出てきた行を採用
            End If
        End If
    Next i

    ' ---- 転記先を走査して書き込む ----
    dstLast = dstWs.Cells(dstWs.Rows.Count, dstKey.Column).End(xlUp).Row
    If dstLast < dstKey.Row Then
        MsgBox "転記先のキーの列（" & ColLetter(dstKey.Column) & "列）にデータがありません。", vbExclamation, TITLE_TEXT
        Exit Sub
    End If
    dstKeys = ReadCol(dstWs, dstKey.Column, dstKey.Row, dstLast)

    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    appChanged = True

    For i = 1 To UBound(dstKeys, 1)
        r = dstKey.Row + i - 1
        k = NormKey(dstKeys(i, 1))
        If Len(k) = 0 Then
            nBlankKey = nBlankKey + 1
        Else
            srcIdx = LookupIndex(keyIndex, k)
            If srcIdx = 0 Then
                nUnmatched = nUnmatched + 1
                If MARK_UNMATCHED Then dstWs.Cells(r, dstKey.Column).Interior.Color = vbYellow
            Else
                nMatched = nMatched + 1
                used(srcIdx) = True
                ' 一致した行のセルだけ書き換える（不一致行の既存値・数式には触れない）
                For m = 1 To nMap
                    v = srcData(m)(srcIdx, 1)
                    If Not (SKIP_BLANK_SOURCE And IsBlankValue(v)) Then
                        dstWs.Cells(r, dstCols(m)).Value = KeepText(v)
                    End If
                Next m
            End If
        End If
    Next i

    For i = 1 To UBound(srcKeys, 1)
        If isFirst(i) And Not used(i) Then nUnused = nUnused + 1
    Next i

Finish:
    errNum = Err.Number
    errDesc = Err.Description
    If appChanged Then
        Application.Calculation = prevCalc
        Application.ScreenUpdating = True
    End If
    If errNum <> 0 Then
        MsgBox "転記の途中でエラーが発生しました: " & errDesc & vbCrLf & _
               "（途中まで書き込まれています。転記先は保存せずに閉じれば元に戻ります）", vbCritical, TITLE_TEXT
        Exit Sub
    End If

    MsgBox "転記が完了しました。" & vbCrLf & vbCrLf & _
           "転記した行数: " & nMatched & vbCrLf & _
           "転記元にキーが無かった行: " & nUnmatched & IIf(MARK_UNMATCHED And nUnmatched > 0, "（キーのセルを黄色にしました）", "") & vbCrLf & _
           "転記先のキーが空欄の行: " & nBlankKey & vbCrLf & _
           "転記元の重複キー: " & nDup & IIf(nDup > 0, "（上にある行を使いました）", "") & vbCrLf & _
           "転記先で使われなかった転記元のキー: " & nUnused & vbCrLf & vbCrLf & _
           "※ 転記先ファイルはまだ保存していません。確認してから保存してください。", _
           vbInformation, TITLE_TEXT
End Sub

'------------------------------------------------------------
' 前回の指定の保存・読み込み（シート名・列・開始行。ファイル名は毎回選ぶ）
'------------------------------------------------------------
Private Sub SaveSettings(ByVal srcKey As Range, ByVal dstKey As Range, _
                         ByRef srcCols() As Long, ByRef dstCols() As Long, ByVal n As Long)
    Dim i As Long, spec As String
    For i = 1 To n
        spec = spec & IIf(i > 1, ",", "") & srcCols(i) & ">" & dstCols(i)
    Next i
    SaveSetting REG_APP, REG_SECTION, "SrcSheet", srcKey.Worksheet.Name
    SaveSetting REG_APP, REG_SECTION, "SrcCell", srcKey.Address(False, False)
    SaveSetting REG_APP, REG_SECTION, "DstSheet", dstKey.Worksheet.Name
    SaveSetting REG_APP, REG_SECTION, "DstCell", dstKey.Address(False, False)
    SaveSetting REG_APP, REG_SECTION, "Map", spec
End Sub

' 前回の指定があり、選んだファイルに同じ名前のシートがあれば True
Private Function LoadSettings(ByVal srcWb As Workbook, ByVal dstWb As Workbook, _
                              ByRef srcKey As Range, ByRef dstKey As Range, _
                              ByRef srcCols() As Long, ByRef dstCols() As Long, ByRef n As Long) As Boolean
    Dim spec As String, pairs() As String, parts() As String, i As Long
    On Error GoTo NoSettings
    spec = GetSetting(REG_APP, REG_SECTION, "Map", "")
    If Len(spec) = 0 Then Exit Function
    Set srcKey = srcWb.Worksheets(GetSetting(REG_APP, REG_SECTION, "SrcSheet", "")) _
                     .Range(GetSetting(REG_APP, REG_SECTION, "SrcCell", ""))
    Set dstKey = dstWb.Worksheets(GetSetting(REG_APP, REG_SECTION, "DstSheet", "")) _
                     .Range(GetSetting(REG_APP, REG_SECTION, "DstCell", ""))
    pairs = Split(spec, ",")
    n = UBound(pairs) + 1
    ReDim srcCols(1 To n)
    ReDim dstCols(1 To n)
    For i = 1 To n
        parts = Split(pairs(i - 1), ">")
        srcCols(i) = CLng(parts(0))
        dstCols(i) = CLng(parts(1))
    Next i
    LoadSettings = True
    Exit Function
NoSettings:
    Set srcKey = Nothing
    Set dstKey = Nothing
    n = 0
End Function

'------------------------------------------------------------
' 表示用
'------------------------------------------------------------
Private Function Describe(ByVal srcKey As Range, ByVal dstKey As Range, _
                          ByRef srcCols() As Long, ByRef dstCols() As Long, ByVal n As Long) As String
    Dim s As String
    s = "■ 転記元: " & srcKey.Worksheet.Parent.Name & " ／ シート「" & srcKey.Worksheet.Name & "」" & vbCrLf & _
        "　 キー: " & ColLabel(srcKey.Worksheet, srcKey.Column, srcKey.Row) & "（" & srcKey.Row & "行目から）" & vbCrLf & _
        "■ 転記先: " & dstKey.Worksheet.Parent.Name & " ／ シート「" & dstKey.Worksheet.Name & "」" & vbCrLf & _
        "　 キー: " & ColLabel(dstKey.Worksheet, dstKey.Column, dstKey.Row) & "（" & dstKey.Row & "行目から）" & vbCrLf & _
        "■ 転記する列" & vbCrLf & _
        PairList(srcKey, dstKey, srcCols, dstCols, n)
    Describe = s
End Function

Private Function PairList(ByVal srcKey As Range, ByVal dstKey As Range, _
                          ByRef srcCols() As Long, ByRef dstCols() As Long, ByVal n As Long) As String
    Dim i As Long, s As String
    If n = 0 Then Exit Function
    s = ""
    For i = 1 To n
        s = s & "　 " & ColLabel(srcKey.Worksheet, srcCols(i), srcKey.Row) & _
                " → " & ColLabel(dstKey.Worksheet, dstCols(i), dstKey.Row) & vbCrLf
    Next i
    PairList = s
End Function

' 「B列『単価』」のように、列記号と見出し（データ開始行の1つ上のセル）を返す
Private Function ColLabel(ByVal ws As Worksheet, ByVal col As Long, ByVal firstRow As Long) As String
    Dim h As String
    If firstRow > 1 Then h = Trim$(ws.Cells(firstRow - 1, col).Text)
    If Len(h) > 20 Then h = Left$(h, 20) & "…"
    ColLabel = ColLetter(col) & "列" & IIf(Len(h) > 0, "『" & h & "』", "")
End Function

Private Function ColLetter(ByVal col As Long) As String
    Dim s As String
    Do While col > 0
        s = Chr$(65 + (col - 1) Mod 26) & s
        col = (col - 1) \ 26
    Loop
    ColLetter = s
End Function

'------------------------------------------------------------
' 小物
'------------------------------------------------------------
' 1 列分を 2 次元配列 (1 To 行数, 1 To 1) で返す（1 行だけでも配列になるようにする）
Private Function ReadCol(ByVal ws As Worksheet, ByVal col As Long, ByVal r1 As Long, ByVal r2 As Long) As Variant
    Dim a() As Variant
    If r1 = r2 Then
        ReDim a(1 To 1, 1 To 1)
        a(1, 1) = ws.Cells(r1, col).Value
        ReadCol = a
    Else
        ReadCol = ws.Range(ws.Cells(r1, col), ws.Cells(r2, col)).Value
    End If
End Function

' 照合用にキーを正規化する
Private Function NormKey(ByVal v As Variant) As String
    Dim s As String
    If IsError(v) Or IsEmpty(v) Then Exit Function
    s = CStr(v)
    If NORMALIZE_WIDTH Then s = StrConv(s, vbNarrow)
    NormKey = Trim$(s)
End Function

' 索引にあれば転記元の行インデックス、無ければ 0
Private Function LookupIndex(ByVal idx As Collection, ByVal k As String) As Long
    On Error Resume Next
    LookupIndex = idx(k)
    On Error GoTo 0
End Function

' 文字列 "001" や "2026/1/1" が数値・日付に化けないよう、必要なときだけ ' を前置する
Private Function KeepText(ByVal v As Variant) As Variant
    If VarType(v) = vbString Then
        If Len(v) > 0 Then
            If IsNumeric(v) Or IsDate(v) Or Left$(v, 1) = "=" Then v = "'" & v
        End If
    End If
    KeepText = v
End Function

Private Function IsBlankValue(ByVal v As Variant) As Boolean
    If IsError(v) Then Exit Function
    If IsEmpty(v) Then
        IsBlankValue = True
    ElseIf VarType(v) = vbString Then
        IsBlankValue = (Len(v) = 0)
    End If
End Function
