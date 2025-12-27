Option Explicit

' =============================================================================
' 事故らない Concat / DeConcat（標準モジュール .bas 用）
'
' ✅ 安全性の要点
'   - Selection が Range でない / 複数範囲選択 を拒否
'   - エラー値(#N/A等)・Null・Empty・数値混在でも落ちない
'   - スペース正規化（全角/タブ/改行→半角、連続スペース潰し）
'   - Concat は原則「右列を消さない」(任意で二段確認して削除)
'   - DeConcat は「右隣列に値があれば止める」(任意で確認して上書き)
'   - Excel状態（計算/イベント/画面更新/警告）を確実に復元
'
' ★ GitHubに置くなら: ConcatTools.bas などで保存が自然
' =============================================================================

' --- 公開マクロ ---------------------------------------------------------------

Public Sub ConcatUltimate_Safe()
    Dim st As AppState
    st = FreezeExcel()

    On Error GoTo EH

    Dim rng As Range
    If Not TryGetSingleAreaSelection(rng) Then GoTo CleanExit

    If rng.Columns.Count <> 2 Then
        MsgBox "二列を選択してください（例: A:B のように隣接した2列）。", vbExclamation
        GoTo CleanExit
    End If

    If rng.Columns(2).Column <> rng.Columns(1).Column + 1 Then
        MsgBox "隣接した2列を選択してください（例: A:B）。", vbExclamation
        GoTo CleanExit
    End If

    Dim arr As Variant
    arr = rng.Value2

    Dim n As Long
    n = UBound(arr, 1)

    Dim outArr() As Variant
    ReDim outArr(1 To n, 1 To 1)

    Dim r As Long
    Dim blankB As Long, badCells As Long

    For r = 1 To n
        Dim leftVal As String, rightVal As String

        leftVal = ToSafeText(arr(r, 1), badCells)
        rightVal = ToSafeText(arr(r, 2), badCells)

        leftVal = Trim$(NormalizeSpaces(leftVal))
        rightVal = Trim$(NormalizeSpaces(rightVal))

        If Len(rightVal) = 0 Then
            blankB = blankB + 1
            outArr(r, 1) = leftVal
        Else
            outArr(r, 1) = leftVal & " " & rightVal
        End If
    Next r

    ' 破壊は最小に：左列だけ上書き、右列は残す（事故防止）
    rng.Columns(1).Value2 = outArr

    Dim msg As String
    msg = "【Concat 完了】" & vbCrLf & _
          "処理行数: " & n & vbCrLf & _
          "右列空白: " & blankB & vbCrLf & _
          "非文字/エラー等: " & badCells & vbCrLf & vbCrLf & _
          "右列（2列目）を削除しますか？（原則は削除しません）"

    If MsgBox(msg, vbQuestion Or vbYesNo) = vbYes Then
        If MsgBox("本当に右列を削除します。シート全体が左詰めされます。続行しますか？", _
                  vbExclamation Or vbYesNo) = vbYes Then
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


Public Sub DeConcatUltimate_Safe()
    Dim st As AppState
    st = FreezeExcel()

    On Error GoTo EH

    Dim rng As Range
    If Not TryGetSingleAreaSelection(rng) Then GoTo CleanExit

    If rng.Columns.Count <> 1 Then
        MsgBox "一列だけを選択してください。", vbExclamation
        GoTo CleanExit
    End If

    ' 右端列だと出力できない
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

    Dim dst As Range
    Set dst = rng.Offset(0, 1).Resize(n, 1)

    ' 右隣に値があるなら基本停止（事故防止）
    If HasAnyValue(dst) Then
        If MsgBox("右隣列に既にデータがあります。上書きして続行しますか？", vbExclamation Or vbYesNo) <> vbYes Then
            GoTo CleanExit
        End If
    End If

    Dim r As Long
    Dim badCells As Long

    For r = 1 To n
        Dim s As String
        s = ToSafeText(arr(r, 1), badCells)
        s = Trim$(NormalizeSpaces(s))

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

    rng.Value2 = leftArr
    dst.Value2 = rightArr

    MsgBox "【DeConcat 完了】" & vbCrLf & _
           "処理行数: " & n & vbCrLf & _
           "非文字/エラー等: " & badCells, vbInformation

CleanExit:
    RestoreExcel st
    Exit Sub

EH:
    MsgBox "予期せぬエラー: " & Err.Description, vbCritical
    Resume CleanExit
End Sub


' --- Excel状態の安全な切替 ----------------------------------------------------

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


' --- Selection 安全化 ---------------------------------------------------------

Private Function TryGetSingleAreaSelection(ByRef rng As Range) As Boolean
    On Error GoTo EH

    If TypeName(Selection) <> "Range" Then
        MsgBox "セル範囲を選択してください。", vbExclamation
        TryGetSingleAreaSelection = False
        Exit Function
    End If

    Set rng = Selection

    If rng.Areas.Count <> 1 Then
        MsgBox "飛び飛び選択（Ctrlで複数範囲）は事故りやすいので非対応です。連続範囲を選択してください。", vbExclamation
        TryGetSingleAreaSelection = False
        Exit Function
    End If

    TryGetSingleAreaSelection = True
    Exit Function

EH:
    MsgBox "選択範囲取得でエラー: " & Err.Description, vbCritical
    TryGetSingleAreaSelection = False
End Function

Private Function HasAnyValue(ByVal rng As Range) As Boolean
    On Error GoTo EH
    HasAnyValue = (Application.WorksheetFunction.CountA(rng) > 0)
    Exit Function
EH:
    ' 判定に失敗したら安全側（ある扱い）に倒す
    HasAnyValue = True
End Function


' --- 文字列化・正規化 ---------------------------------------------------------

Private Function ToSafeText(ByVal v As Variant, ByRef badCount As Long) As String
    ' エラー値/Null/Empty を吸収し、絶対に落ちない文字列化
    On Error GoTo EH

    If IsError(v) Then
        badCount = badCount + 1
        ToSafeText = vbNullString
        Exit Function
    End If

    If IsNull(v) Or IsEmpty(v) Then
        ToSafeText = vbNullString
        Exit Function
    End If

    ToSafeText = CStr(v)
    Exit Function

EH:
    badCount = badCount + 1
    ToSafeText = vbNullString
End Function

Private Function NormalizeSpaces(ByVal s As String) As String
    ' 全角/タブ/改行を半角スペースへ寄せ、連続スペースを1個に潰す
    Dim t As String
    t = s

    t = Replace(t, vbTab, " ")
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    t = Replace(t, "　", " ")

    On Error Resume Next
    t = Application.WorksheetFunction.Trim(t)  ' 連続スペース潰し + Trim
    On Error GoTo 0

    NormalizeSpaces = t
End Function
