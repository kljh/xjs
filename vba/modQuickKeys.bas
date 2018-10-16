Attribute VB_Name = "modQuickKeys"
Option Explicit
' Author : Claude Cochet
' Created : 2010-09-17

Sub QuickKeys()

    'SaveSetting "Quick", "Keys", "Enable", True
    'If Not CBool(GetSetting("Quick", "Keys", "Enable", False)) Then Exit Sub

    ' Reminder : SHIFT is +, CTRL is ^ (caret), ALT is %, ALPHA ENTER is ~ or {RETURN}, NUMERIC KEYPAD ENTER is {ENTER}
    ' Other keys: {UP}, {DOWN}, {LEFT}, {RIGHT}, {PGUP}, {PGDN}, {DOWN}, {ESC},  {DEL}, {CLEAR}, {BS}, {BREAK}, {HOME}, {INS}, {TAB}, {F1},.. {F15}
    Application.OnKey "{F1}", "SendKeysF2"          ' F1 is useless, time wasting
    Application.OnKey "^!", "ToggleTwoDigits"       ' extend standard shortcut
    Application.OnKey "^{%}", "TogglePercent"       ' extend standard shortcut
    Application.OnKey "^#", "ToggleDateFormat"      ' extend standard shortcut
    Application.OnKey "^{~}", "RemoveFormat"        ' extend standard shortcut
    Application.OnKey "^A", "AdjustRange"           ' Ctrl+Shift+A auto adjust formula array
    Application.OnKey "^T", "TransposeRange"        ' Ctrl+Shift+T transpose formula

    Assistant.KeyboardShortcutTips = True
End Sub

Private Sub SendKeysF2()
    Application.SendKeys ("{F2}")
End Sub

Private Sub ToggleTwoDigits()
    Dim rng As Range, fmt As String
    Set rng = Selection
    fmt = rng.Cells(1, 1).NumberFormat

    If fmt = "#,##0.00" Then
        rng.NumberFormat = "#,##0"
    ElseIf fmt = "#,##0" Then
        rng.NumberFormat = "#,##0.0000"
    Else
        rng.NumberFormat = "#,##0.00"
    End If
End Sub

Private Sub TogglePercent()
    Dim rng As Range, fmt As String
    Set rng = Selection
    fmt = rng.Cells(1, 1).NumberFormat

    If fmt = "0%" Then
        rng.NumberFormat = "0.00%"
    ElseIf fmt = "0.00%" Then
        rng.NumberFormat = "0.0000%"
    Else
        rng.NumberFormat = "0%"
    End If
End Sub

Private Sub ToggleDateFormat()
    Dim rng As Range, fmt As String
    Set rng = Selection
    fmt = rng.Cells(1, 1).NumberFormat

    If fmt = "d-mmm-yy" Then
        rng.NumberFormat = "ddd dd-mmm-yyyy"
    ElseIf fmt = "ddd dd-mmm-yyyy" Then
        rng.NumberFormat = "ddd dd-mmm-yyyy hh:mm"
    Else
        rng.NumberFormat = "d-mmm-yy"
    End If
End Sub

Private Sub RemoveFormat()
    Dim rng As Range, fmt As String
    Set rng = Selection
    fmt = rng.Cells(1, 1).NumberFormat

    If fmt = "General" Then
        rng.HorizontalAlignment = xlGeneral
    Else
        rng.NumberFormat = "General"
    End If
End Sub

' Adjust a selected Range array to the correct size
Public Sub AdjustRange()

    Dim rng As Range, rng_adjusted As Range
    Dim formula_txt As String

    On Error Resume Next
    ' formula
    Set rng = ActiveCell
    formula_txt = rng.Formula
    ' formula array
    Set rng = ActiveCell.CurrentArray
    formula_txt = rng.FormulaArray
    On Error GoTo 0

    ' Evaluate the formula (Excel4 macro)
    Dim evaluation As Variant
    evaluation = ActiveSheet.Evaluate(formula_txt)

    ' If the result is an error, then do nothing
    If IsError(evaluation) Then Exit Sub

    ' Check the target Range fits the s/s
    Dim nb_rows As Long, nb_cols As Long: nb_rows = 1: nb_cols = 0
    On Error Resume Next
    nb_rows = UBound(evaluation, 1)
    nb_cols = UBound(evaluation, 2)
    If nb_cols = 0 Then
        ' single row 2d arrays are forced into 1D array
        nb_cols = nb_rows
        nb_rows = 1
    End If
    Set rng_adjusted = rng.Resize(nb_rows, nb_cols)
    On Error GoTo 0
    If rng_adjusted Is Nothing Then
        MsgBox "Target Range does not fit the s/s. Rows x cols: " & nb_rows & " x " & nb_cols
        Exit Sub
    End If

    ' Check whether we're done
    If rng.Address = rng_adjusted.Address Then Exit Sub

    ' Check that extra destination cells are empty
    Dim nb_inter As Long, nb_adjust As Long
    nb_inter = Application.WorksheetFunction.CountA(Intersect(rng, rng_adjusted))
    nb_adjust = Application.WorksheetFunction.CountA(rng_adjusted)
    If nb_adjust > nb_inter Then
        rng_adjusted.Select
        MsgBox "Target Range is not empty"
        Exit Sub
    End If

    ' Adjust the formula array Range
    rng.ClearContents
    rng_adjusted.FormulaArray = formula_txt
    rng_adjusted.Select
End Sub

Public Sub TransposeRange()
    Dim rng As Range
    Set rng = ActiveCell.CurrentArray

    Dim formula_txt As String, new_formula_txt As String, transpose_txt As String
    formula_txt = rng.FormulaArray
    transpose_txt = "=TRANSPOSE("
    If Left(formula_txt, Len(transpose_txt)) = transpose_txt Then
        new_formula_txt = "=" & Mid(formula_txt, Len(transpose_txt) + 1, Len(formula_txt) - Len(transpose_txt) - 1)
    Else
        new_formula_txt = transpose_txt & Mid(formula_txt, 2) & ")"
    End If

    rng.FormulaArray = new_formula_txt
End Sub
