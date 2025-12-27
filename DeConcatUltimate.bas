Option Explicit

Sub DeConcatUltimate()
    Dim rng As Range, arr, leftArr(), rightArr()
    Dim r As Long
    Dim firstToken As String, remainder As String
    Dim pos As Long, errors As Long
    
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False
    
    On Error GoTo CleanFail
    
    Set rng = Selection
    If rng.Columns.Count <> 1 Then
        MsgBox "一列だけを選択してください。"
        GoTo CleanExit
    End If
    
    arr = rng.Value
    ReDim leftArr(1 To UBound(arr, 1), 1 To 1)
    ReDim rightArr(1 To UBound(arr, 1), 1 To 1)
    
    For r = 1 To UBound(arr, 1)
        On Error Resume Next
        
        Dim s As String
        s = NormalizeSpaces(arr(r, 1))
        
        pos = InStr(s, " ")
        If pos > 0 Then
            firstToken = Left$(s, pos - 1)
            remainder = Mid$(s, pos + 1)
        Else
            firstToken = s
            remainder = ""
        End If
        
        leftArr(r, 1) = firstToken
        rightArr(r, 1) = remainder
        
        If Err.Number <> 0 Then
            errors = errors + 1
            Err.Clear
        End If
        On Error GoTo 0
    Next r
    
    rng.Offset(0, 1).Resize(UBound(arr, 1), 1).Insert Shift:=xlShiftToRight
    rng.Value = leftArr
    rng.Offset(0, 1).Value = rightArr
    
    MsgBox "【DeConcat 完了】" & vbCrLf & _
           "処理行数: " & UBound(arr, 1) & vbCrLf & _
           "エラー行: " & errors
    
    GoTo CleanExit
    
CleanFail:
    MsgBox "予期せぬエラー: " & Err.Description
    
CleanExit:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.EnableEvents = True
End Sub
