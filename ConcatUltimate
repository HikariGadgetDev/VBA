Option Explicit

Sub ConcatUltimate()
    Dim rng As Range, arr, result()
    Dim r As Long, last As Long
    Dim col As Long, rowStart As Long
    Dim blankB As Long, errors As Long
    Dim threshold As Long
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    On Error GoTo CleanFail
    
    Set rng = Selection
    If rng.Columns.Count <> 2 Then
        MsgBox "二列を選択してください。"
        GoTo CleanExit
    End If
    
    col = rng.Column
    rowStart = rng.Row
    last = rng.Row + rng.Rows.Count - 1
    
    arr = rng.Value
    ReDim result(1 To UBound(arr, 1), 1 To 1)
    
    threshold = 200000 ' 20万セル以上は高速ルート前提
    
    For r = 1 To UBound(arr, 1)
        On Error Resume Next
        
        Dim leftVal As String, rightVal As String
        
        leftVal = Trim(SanitizeText(arr(r, 1)))
        rightVal = Trim(SanitizeText(arr(r, 2)))
        
        If Len(rightVal) = 0 Then
            blankB = blankB + 1
            result(r, 1) = leftVal
        Else
            result(r, 1) = leftVal & " " & rightVal
        End If
        
        If Err.Number <> 0 Then
            errors = errors + 1
            Err.Clear
        End If
        On Error GoTo 0
    Next r
    
    rng.Columns(1).Value = result
    rng.Columns(2).Delete Shift:=xlToLeft

    MsgBox "【Concat 完了】" & vbCrLf & _
           "結合行数: " & UBound(arr, 1) & vbCrLf & _
           "B列空白: " & blankB & vbCrLf & _
           "エラー行: " & errors
    
    GoTo CleanExit

CleanFail:
    MsgBox "処理中に予期せぬエラー: " & Err.Description

CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
