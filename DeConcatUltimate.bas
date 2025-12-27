Option Explicit

' ============================================================
'   - エラー値/Null/Empty/数値の混在に耐える
'   - 複数範囲選択は拒否（事故源）
'   - 右隣にデータがある場合は確認 or 中止
'   - 原則として列削除をしない（必要時のみ確認して削除）
' ============================================================

Public Sub ConcatSafe()
    Dim st As AppState
    st = FreezeExcel()

    On Error GoTo EH

    Dim rng As Range
    If Not TryGetSingleAreaSelection(rng) Then GoTo CleanExit

    If rng.Columns.Count <> 2 Then
        MsgBox "二列（隣接した2列）を選択してください。", vbExclamation
        GoTo CleanExit
    End If

    ' なるべく「隣接2列」前提を明確化（事故防止）
    If rng.Columns(2).Column <> rng.Columns(1).Column + 1 Then
        MsgBox "隣接した2列を選択してください。（例：A:B）", vbExclamation
        GoTo CleanExit
    End If

    Dim arr As Variant
    arr = rng.Value2  ' Value2で余計な日付変換などを減らす

    Dim n As Long
    n = UBound(arr, 1)

    Dim outArr() As Variant
    ReDim outArr(1 To n, 1 To 1)

    Dim r As Long
    Dim blankB As Long, errCells As Long

    For r = 1 To n
        Dim leftVal As String, rightVal As String
        leftVal = SanitizeText(arr(r, 1), errCells)
        rightVal = SanitizeText(arr(r, 2), errCells)

        leftVal = Trim$(NormalizeSpaces(leftVal))
        rightVal = Trim$(NormalizeSpaces(rightVal))

        If Len(rightVal) = 0 Then
            blankB = blankB + 1
            outArr(r, 1) = leftVal
        Else
            outArr(r, 1) = leftVal & " " & rightVal
        End If
    Next r

    ' 出力：A列に上書き（B列は原則残す）
    rng.Columns(1).Value2 = outArr

    Dim msg As String
    msg = "【Concat 完了】" & vbCrLf & _
          "処理行数: " & n & vbCrLf & _
          "B列空白: " & blankB & vbCrLf & _
          "エラー/非文字セル: " & errCells & vbCrLf & vbCrLf & _
          "B列を削除しますか？（原則は削除しない）"

    If MsgBox(msg, vbQuestion Or vbYesNo) = vbYes Then
        ' B列削除は破壊的なので、さらに安全確認
        If MsgBox("本当に2列目（右列）を削除します。元に戻せません。続行？", vbExclamation Or vbYesNo) = vbYes Then
            rng.Columns(2).Delete Shift:=xlToLeft
        End If
    End If

CleanExit:
    RestoreExcel st
    Exit Sub

EH:
    MsgBox "処理中に予期せぬエラー: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

Public Sub DeConcatSafe()
    Dim st As AppState
    st = FreezeExcel()

    On Error GoTo EH

    Dim rng As Range
    If Not TryGetSingleAreaSelection(rng) Then GoTo CleanExit

    If rng.Columns.Count <> 1 Then
        MsgBox "一列だけを選択してください。", vbExclamation
        GoTo CleanExit
    End If

    ' 右端列での挿入事故を防止
    If rng.Column >= rng.Worksheet.Columns.Count Then
        MsgBox "右隣の列が存在しないため分割できません。", vbExclamation
        GoTo CleanExit
    End If

    Dim arr As Variant
    arr = rng.Value2

    Dim n As Long
    n = UBound(arr, 1)

    Dim leftArr() As Variant, rightArr() As Variant
    ReDim leftArr(1 To n, 1 To 1)
    ReDim rightArr(1 To n, 1 To 1)

    ' 出力先（右列）が埋まってたら止める（事故防止）
    Dim dst As Range
    Set dst = rng.Offset(0, 1).Resize(n, 1)

    If HasAnyValue(dst) Then
        If MsgBox("右隣列に既にデータがあります。上書きすると壊れます。上書きして続行しますか？", _
                  vbExclamation Or vbYesNo) <> vbYes Then
            GoTo CleanExit
        End If
    End If

    Dim r As Long
    Dim errCells As Long

    For r = 1 To n
        Dim s As String
        s = SanitizeText(arr(r, 1), errCells)
        s = NormalizeSpaces(s)
        s = Trim$(s)

        Dim pos As Long
        pos = InStr(1, s, " ", vbBinaryCompare)

        If pos > 0 Then
            leftArr(r, 1) = Left$(s, pos - 1)
            rightArr(r, 1) = Mid$(s, pos + 1)
        Else
            leftArr(r, 1) = s
            rightArr(r, 1) = vbNullString
        End If
    Next r

    ' 出力：元列を左、右隣へ残り
    rng.Value2 = leftArr
    dst.Value2 = rightArr

    MsgBox "【DeConcat 完了】" & vbCrLf & _
           "処理行数: " & n & vbCrLf & _
           "エラー/非文字セル: " & errCells, vbInformation

CleanExit:
    RestoreExcel st
    Exit Sub

EH:
    MsgBox "予期せぬエラー: " & Err.Description, vbCritical
    Resume CleanExit
End Sub

' --- Helpers (Safety) ---------------------------------------------------------

Private Type AppState
    ScreenUpdating As Boolean
    EnableEvents As Boolean
    CalcMode As XlCalculation
    DisplayAlerts As Boolean
End Type

Private Function FreezeExcel() As AppState
    With FreezeExcel
        .ScreenUpdating = Application.ScreenUpdating
        .EnableEvents = Application.EnableEvents
        .CalcMode = Application.Calculation
        .DisplayAlerts = Application.DisplayAlerts
    End With

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
End Function

Private Sub RestoreExcel(ByVal st As AppState)
    Application.ScreenUpdating = st.ScreenUpdating
    Application.EnableEvents = st.EnableEvents
    Application.Calculation = st.CalcMode
    Application.DisplayAlerts = st.DisplayAlerts
End Sub

Private Function TryGetSingleAreaSelection(ByRef rng As Range) As Boolean
    On Error GoTo EH

    If TypeName(Selection) <> "Range" Then
        MsgBox "セル範囲を選択してください。", vbExclamation
        TryGetSingleAreaSelection = False
        Exit Function
    End If

    Set rng = Selection

    If rng.Areas.Count <> 1 Then
        MsgBox "複数範囲選択（Ctrlで飛び飛び選択）は事故りやすいので非対応です。1つの連続範囲を選択してください。", vbExclamation
        TryGetSingleAreaSelection = False
        Exit Function
    End If

    TryGetSingleAreaSelection = True
    Exit Function

EH:
    MsgBox "選択範囲の取得でエラー: " & Err.Description, vbCritical
    TryGetSingleAreaSelection = False
End Function

Private Function HasAnyValue(ByVal rng As Range) As Boolean
    ' 速い空判定：CountAはエラー値もカウントされる
    On Error GoTo EH
    HasAnyValue = (Application.WorksheetFunction.CountA(rng) > 0)
    Exit Function
EH:
    ' CountAが落ちることは稀だが、落ちたら安全側で「ある」と見なす
    HasAnyValue = True
End Function

' --- Text Normalization -------------------------------------------------------

Private Function SanitizeText(ByVal v As Variant, ByRef errCount As Long) As String
    ' エラー値やNull等を安全に文字列化し、事故を防ぐ
    On Error GoTo EH

    If IsError(v) Then
        errCount = errCount + 1
        SanitizeText = vbNullString
        Exit Function
    End If

    If IsNull(v) Or IsEmpty(v) Then
        SanitizeText = vbNullString
        Exit Function
    End If

    ' 文字列化（数値/日付等もここで吸収）
    SanitizeText = CStr(v)
    Exit Function

EH:
    errCount = errCount + 1
    SanitizeText = vbNullString
End Function

Private Function NormalizeSpaces(ByVal s As String) As String
    ' 全角スペース/タブ/改行を半角スペースに寄せ、連続スペースを1個に潰す
    Dim t As String
    t = s

    t = Replace(t, vbTab, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    t = Replace(t, "　", " ") ' 全角スペース

    ' 連続スペースを潰す（WorksheetFunction.Trimは連続スペースを1つにする）
    On Error Resume Next
    t = Application.WorksheetFunction.Trim(t)
    On Error GoTo 0

    NormalizeSpaces = t
End Function
