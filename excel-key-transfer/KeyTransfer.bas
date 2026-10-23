Option Explicit

'============================================================
' キー一致転記マクロ
'   転記元の「キー列」と転記先の「キー列」が一致する行について、
'   転記元の指定列の値を、転記先の指定列へ書き込む。
'   （VLOOKUP / XLOOKUP を値で貼り付けるのと同じ結果）
'
' 使い方：下の「設定」を書き換えて TransferByKey を実行。
'         ファイル選択ダイアログが 2 回出る（1 回目=転記先、2 回目=転記元）。
'============================================================

'==================== 設定 ====================
' --- 転記元 ---
Private Const SRC_SHEET As String = "Sheet1"   ' シート名
Private Const SRC_KEY_COL As String = "A"      ' キー列（列記号）
Private Const SRC_FIRST_ROW As Long = 2        ' データ開始行（見出しの次の行）

' --- 転記先 ---
Private Const DST_SHEET As String = "Sheet1"
Private Const DST_KEY_COL As String = "C"
Private Const DST_FIRST_ROW As Long = 2

' --- 転記する列の対応 ---
' 「転記元の列>転記先の列」をカンマ区切りで並べる。何組でも可。
' 例 "B>E, D>F" … 転記元B列→転記先E列、転記元D列→転記先F列
Private Const COLUMN_MAP As String = "B>E, D>F"

' --- オプション ---
Private Const SKIP_BLANK_SOURCE As Boolean = False ' True: 転記元が空欄なら転記先を上書きしない
Private Const NORMALIZE_WIDTH As Boolean = True    ' True: キーの全角/半角・前後の空白を無視して照合
Private Const MARK_UNMATCHED As Boolean = True     ' True: 転記元に見つからなかった転記先のキーセルを黄色にする
'==============================================

Public Sub TransferByKey()
    Dim srcWb As Workbook, dstWb As Workbook
    Dim srcOpenedHere As Boolean, dstOpenedHere As Boolean
    Dim srcWs As Worksheet, dstWs As Worksheet
    Dim srcKeyCol As Long, dstKeyCol As Long
    Dim srcCols() As Long, dstCols() As Long, nMap As Long
    Dim srcLast As Long, dstLast As Long
    Dim srcKeys As Variant, dstKeys As Variant, srcData() As Variant
    Dim keyIndex As Collection, isFirst() As Boolean, used() As Boolean
    Dim i As Long, m As Long, r As Long, srcIdx As Long
    Dim k As String, v As Variant
    Dim nMatched As Long, nUnmatched As Long, nBlankKey As Long, nDup As Long, nUnused As Long
    Dim appChanged As Boolean, prevCalc As XlCalculation

    On Error GoTo Fail

    ' 転記先は書き込むので通常モード、転記元は読み取り専用で開く
    Set dstWb = PickWorkbook("【1/2】転記先のファイルを選択", False, dstOpenedHere)
    If dstWb Is Nothing Then Exit Sub
    Set srcWb = PickWorkbook("【2/2】転記元のファイルを選択", True, srcOpenedHere)
    If srcWb Is Nothing Then Exit Sub

    Set srcWs = GetSheet(srcWb, SRC_SHEET)
    Set dstWs = GetSheet(dstWb, DST_SHEET)
    If srcWs Is Nothing Or dstWs Is Nothing Then GoTo Cleanup

    srcKeyCol = ColNum(SRC_KEY_COL)
    dstKeyCol = ColNum(DST_KEY_COL)
    If srcKeyCol = 0 Or dstKeyCol = 0 Then
        MsgBox "キー列の指定が不正です: 転記元=""" & SRC_KEY_COL & """ 転記先=""" & DST_KEY_COL & """", vbExclamation
        GoTo Cleanup
    End If
    If Not ParseMap(COLUMN_MAP, srcCols, dstCols, nMap) Then GoTo Cleanup

    ' ---- 転記元を読み込み、キー → 行番号 の索引を作る ----
    srcLast = srcWs.Cells(srcWs.Rows.Count, srcKeyCol).End(xlUp).Row
    If srcLast < SRC_FIRST_ROW Then
        MsgBox "転記元のキー列（" & SRC_KEY_COL & "列）にデータがありません。", vbExclamation
        GoTo Cleanup
    End If
    srcKeys = ReadCol(srcWs, srcKeyCol, SRC_FIRST_ROW, srcLast)
    ReDim srcData(1 To nMap)
    For m = 1 To nMap
        srcData(m) = ReadCol(srcWs, srcCols(m), SRC_FIRST_ROW, srcLast)
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
    dstLast = dstWs.Cells(dstWs.Rows.Count, dstKeyCol).End(xlUp).Row
    If dstLast < DST_FIRST_ROW Then
        MsgBox "転記先のキー列（" & DST_KEY_COL & "列）にデータがありません。", vbExclamation
        GoTo Cleanup
    End If
    dstKeys = ReadCol(dstWs, dstKeyCol, DST_FIRST_ROW, dstLast)

    prevCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    appChanged = True

    For i = 1 To UBound(dstKeys, 1)
        r = DST_FIRST_ROW + i - 1
        k = NormKey(dstKeys(i, 1))
        If Len(k) = 0 Then
            nBlankKey = nBlankKey + 1
        Else
            srcIdx = LookupIndex(keyIndex, k)
            If srcIdx = 0 Then
                nUnmatched = nUnmatched + 1
                If MARK_UNMATCHED Then dstWs.Cells(r, dstKeyCol).Interior.Color = vbYellow
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

    MsgBox "転記が完了しました。" & vbCrLf & vbCrLf & _
           "転記した行数: " & nMatched & vbCrLf & _
           "転記元にキーが無かった行: " & nUnmatched & IIf(MARK_UNMATCHED And nUnmatched > 0, "（キーセルを黄色表示）", "") & vbCrLf & _
           "転記先のキー空欄行: " & nBlankKey & vbCrLf & _
           "転記元の重複キー: " & nDup & IIf(nDup > 0, "（先に出てくる行を採用）", "") & vbCrLf & _
           "転記先で使われなかった転記元キー: " & nUnused & vbCrLf & vbCrLf & _
           "※ 転記先ファイルはまだ保存していません。確認してから保存してください。", _
           vbInformation, "キー一致転記"

Cleanup:
    On Error Resume Next
    If appChanged Then
        Application.Calculation = prevCalc
        Application.ScreenUpdating = True
    End If
    If srcOpenedHere And Not srcWb Is Nothing Then srcWb.Close SaveChanges:=False
    Exit Sub

Fail:
    MsgBox "エラーが発生しました: " & Err.Description, vbCritical, "キー一致転記"
    Resume Cleanup
End Sub

' ファイル選択。既に開いていればそれを使い、開いていなければ開く。キャンセル時は Nothing。
Private Function PickWorkbook(ByVal title As String, ByVal readOnly As Boolean, ByRef openedHere As Boolean) As Workbook
    Dim f As Variant, wb As Workbook
    openedHere = False
#If Mac Then
    f = Application.GetOpenFilename(Title:=title)
#Else
    f = Application.GetOpenFilename("Excel ファイル (*.xlsx;*.xlsm;*.xls),*.xlsx;*.xlsm;*.xls", , title)
#End If
    If VarType(f) = vbBoolean Then Exit Function

    For Each wb In Application.Workbooks
        If StrComp(wb.FullName, CStr(f), vbTextCompare) = 0 Then
            Set PickWorkbook = wb
            Exit Function
        End If
    Next wb
    Set PickWorkbook = Workbooks.Open(Filename:=CStr(f), ReadOnly:=readOnly)
    openedHere = True
End Function

Private Function GetSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetSheet = wb.Worksheets(sheetName)
    On Error GoTo 0
    If GetSheet Is Nothing Then
        MsgBox "「" & wb.Name & "」にシート「" & sheetName & "」がありません。", vbExclamation
    End If
End Function

' "B>E, D>F" を列番号の配列に分解する
Private Function ParseMap(ByVal spec As String, ByRef srcCols() As Long, ByRef dstCols() As Long, ByRef n As Long) As Boolean
    Dim pairs() As String, parts() As String, i As Long
    pairs = Split(spec, ",")
    n = UBound(pairs) - LBound(pairs) + 1
    If n < 1 Then
        MsgBox "COLUMN_MAP が空です。", vbExclamation
        Exit Function
    End If
    ReDim srcCols(1 To n)
    ReDim dstCols(1 To n)
    For i = 1 To n
        parts = Split(pairs(LBound(pairs) + i - 1), ">")
        If UBound(parts) <> 1 Then
            MsgBox "COLUMN_MAP の書式が不正です: """ & Trim$(pairs(LBound(pairs) + i - 1)) & """" & vbCrLf & _
                   "「転記元列>転記先列」の形で書いてください（例 B>E）。", vbExclamation
            Exit Function
        End If
        srcCols(i) = ColNum(parts(0))
        dstCols(i) = ColNum(parts(1))
        If srcCols(i) = 0 Or dstCols(i) = 0 Then
            MsgBox "COLUMN_MAP の列記号が不正です: """ & Trim$(pairs(LBound(pairs) + i - 1)) & """", vbExclamation
            Exit Function
        End If
    Next i
    ParseMap = True
End Function

' 列記号 → 列番号。不正なら 0
Private Function ColNum(ByVal letters As String) As Long
    Dim i As Long, c As Long
    letters = UCase$(Trim$(letters))
    If Len(letters) = 0 Or Len(letters) > 3 Then Exit Function
    For i = 1 To Len(letters)
        c = Asc(Mid$(letters, i, 1))
        If c < 65 Or c > 90 Then Exit Function
        ColNum = ColNum * 26 + (c - 64)
    Next i
    If ColNum > 16384 Then ColNum = 0
End Function

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
