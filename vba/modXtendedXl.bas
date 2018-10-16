Attribute VB_Name = "modXtendedXl"
' Extended Excel context menu
' Author : Claude Cochet
' Created : 2010-10-06

Option Explicit

#If VBA7 Then
Private Declare PtrSafe Function GetTickCount Lib "kernel32" () As Long
#Else
Private Declare Function GetTickCount Lib "Kernel32.dll" () As Long
#End If

' ********************************************************************************
' Common bar

Private barRightClick As CommandBar
Private barRightClickMarker As String       ' This string is put in all custom button's TooltipText in order for them to be easily recognized for deletion step

Public Sub AddCustomRightClickBar()

    barRightClickMarker = "kTools. "
    RemoveCustomRightClickBar

    ' The following rows uses the default PopUp menu
    Set barRightClick = Application.CommandBars("Cell")
    AddCustomContextBar barRightClick

    AddCustomContextBar Application.CommandBars("Row")
    AddCustomContextBar Application.CommandBars("Column")
End Sub

Private Sub AddCustomContextBar(bar As CommandBar)

    ' Do not use the following accelerator keys (they are used in the default menu):
    ' t/c/p/S : cut/copy/paste/paste special
    ' i/d/n : insert/delete/clearcontents
    ' m : insert comment
    ' f/k/w/h : format/pickfromlist/addwatch/hyperlink
    ' Added :
    ' r/a/L/R/!/T/A : refresh Range / all / left to sheet / sheet to right / selected sheets / display time / activate sheets
    ' x/t : export/import sheets
    ' q : quick mode
    ' y : formula smart display

    ' Control type can be : msoControlPopup, msoControlButton,
    ' msoControlEdit, msoControlComboBox, or msoControlDropdown.
    ' Contrary to a msoControlComboBox control, a msoControlDropDown control can't be edited (ie typed-in).

    With bar
        Dim position As Long: position = .Controls.Count
        ' we use position=.Controls.Count to append at the end
        ' instead of position=1 to append at the beging which is far better (even PW agrees)

        With .Controls.Add(Type:=msoControlPopup, Before:=position, Temporary:=True)
            Dim rndFaceId As Long: rndFaceId = Int(Rnd * 3500 + 1)
            ' .FaceId =
            .Caption = "Too&ls"
            .TooltipText = barRightClickMarker & ""

            Dim innerCmdBar: Set innerCmdBar = .CommandBar
            With innerCmdBar

                With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                    .Style = msoButtonIconAndCaption
                    .FaceId = 1014
                    .Caption = "&Array formula autosize"
                    .OnAction = "'OnFormating """ & .Caption & """ '"
                    .TooltipText = barRightClickMarker & ""
                End With

                With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                    .BeginGroup = True
                    .Style = msoButtonIconAndCaption
                    .FaceId = 107 ' 107 (thunder sheet)
                    .Caption = "E&xport Sheet"
                    .OnAction = "OnExportSheet"
                    .TooltipText = barRightClickMarker & ""
                End With

                With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                    .Style = msoButtonIconAndCaption
                    .FaceId = 2896 ' 2896 (Klaxon, star and text)
                    .Caption = "Impor&t Sheet"
                    .OnAction = "OnImportSheets"
                    .TooltipText = barRightClickMarker & ""
                End With

                With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                    .BeginGroup = True
                    .Style = msoButtonIconAndCaption
                    .FaceId = 1407 ' 1407 (chart)
                    .Caption = "Check Range &Names"
                    .OnAction = "OnCheckRangeNames"
                    .TooltipText = barRightClickMarker & ""
                End With

                With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                    .Style = msoButtonIconAndCaption
                    .FaceId = 1407 ' 1407 (chart)
                    .Caption = "&Reduce FormulaArray"
                    .OnAction = "OnReduceFormulaArray"
                    .TooltipText = barRightClickMarker & ""
                End With

            End With
        End With

        position = position + 1
        With .Controls.Add(Type:=msoControlPopup, Before:=position, Temporary:=True)
            .Caption = "&Refresh"

            Set innerCmdBar = .CommandBar
            With innerCmdBar
                With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                    .BeginGroup = True
                    .Style = msoButtonIconAndCaption
                    .FaceId = rndFaceId  ' 2168 (hand)
                    .Caption = "Ran&ge (" & rndFaceId & ")"
                    .OnAction = "'OnRefresh ""Range"" '"
                    .TooltipText = barRightClickMarker & ""
                End With
                With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                    .Style = msoButtonIconAndCaption
                    .FaceId = 1044 ' 1044 (spline)
                    .Caption = "All Open &Workbooks"
                    .OnAction = "'OnRefresh ""OpenWorkbooks"" '"
                    .TooltipText = barRightClickMarker & ""
                End With
                With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                    .Style = msoButtonIconAndCaption
                    .FaceId = 1044 ' 1044 (spline)
                    .Caption = "Left to Right (&All Sheets)"
                    .OnAction = "'OnRefresh ""LeftToRight"" '"
                    .TooltipText = barRightClickMarker & ""
                End With
                With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                    .Style = msoButtonIconAndCaption
                    .FaceId = 2767 ' 2767 (dog)
                    .Caption = "&Left to Sheet"
                    .OnAction = "'OnRefresh ""LeftToSheet"" '"
                    .TooltipText = barRightClickMarker & ""
                End With
                With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                    .Style = msoButtonIconAndCaption
                    .FaceId = 2767 ' 2767 (dog)
                    .Caption = "Sheet &to Right"
                    .OnAction = "'OnRefresh ""SheetToRight"" '"
                    .TooltipText = barRightClickMarker & ""
                End With
                With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                    .Style = msoButtonIconAndCaption
                    .FaceId = 570 ' 570 (binoculars and an arrow )
                    .Caption = "Selected &Sheets"
                    .OnAction = "'OnRefresh ""SelectedSheets"" '"
                    .TooltipText = barRightClickMarker & ""
                End With
                With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                    .Style = msoButtonCaption
                    .Caption = "Display &Time"      ' do NOT change that name
                    .OnAction = "'OnRefresh ""CalcTimeDisplay"" '"
                    .TooltipText = barRightClickMarker & ""
                End With
                With .Controls.Add(Type:=msoControlDropdown, Temporary:=True)
                    .Caption = "&Activate"          ' do NOT change that name
                    .AddItem "CalculatedSheets", 1
                    .AddItem "CalculatedAndFinal", 2
                    .AddItem "InitialSheet", 3
                    .ListIndex = 2
                    .DropDownWidth = 120
                    ' never called, OnAction has no effect for msoControlDropdown
                    .OnAction = "'OnRefresh ""CalcActivate"" '"
                    .TooltipText = barRightClickMarker & ""
                End With
            End With
        End With

        position = position + 1
        With .Controls.Add(Type:=msoControlPopup, Before:=position, Temporary:=True)
            .Caption = "F&ormating"

            Set innerCmdBar = .CommandBar
            With innerCmdBar
                With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                    .Style = msoButtonIconAndCaption
                    .FaceId = 2063 ' 2063 (dotted square )
                    .Caption = "&Title"
                    .OnAction = "'OnFormating """ & .Caption & """ '"
                    .TooltipText = barRightClickMarker & ""
                End With
                With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                    .Style = msoButtonIconAndCaption
                    .FaceId = 2063
                    .Caption = "&Label"
                    .OnAction = "'OnFormating """ & .Caption & """ '"
                    .TooltipText = barRightClickMarker & ""
                End With
                With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                    .Style = msoButtonIconAndCaption
                    .FaceId = 2063
                    .Caption = "&Input"
                    .OnAction = "'OnFormating """ & .Caption & """ '"
                    .TooltipText = barRightClickMarker & ""
                End With
                With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                    .Style = msoButtonIconAndCaption
                    .FaceId = 2063
                    .Caption = "&Output"
                    .OnAction = "'OnFormating """ & .Caption & """ '"
                    .TooltipText = barRightClickMarker & ""
                End With
                With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                    .Style = msoButtonIconAndCaption
                    .FaceId = 2063
                    .Caption = "&Intermediate"
                    .OnAction = "'OnFormating """ & .Caption & """ '"
                    .TooltipText = barRightClickMarker & ""
                End With
                With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                    .Style = msoButtonIconAndCaption
                    .FaceId = 2063
                    .Caption = "NotRele&vant"
                    .OnAction = "'OnFormating """ & .Caption & """ '"
                    .TooltipText = barRightClickMarker & ""
                End With


                With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                    .BeginGroup = True
                    .Style = msoButtonIconAndCaption
                    .FaceId = 159 ' 159 (multi layer)
                    .Caption = "&Hide #N/A"
                    .OnAction = "'OnFormating """ & .Caption & """ '"
                    .TooltipText = barRightClickMarker & ""
                End With
                With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                    .Style = msoButtonIconAndCaption
                    .FaceId = 159
                    .Caption = "H&ide any #ERR"
                    .OnAction = "'OnFormating """ & .Caption & """ '"
                    .TooltipText = barRightClickMarker & ""
                End With
                With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                    .Style = msoButtonIconAndCaption
                    .FaceId = 159
                    .Caption = "Set as Default"
                    .OnAction = "'OnFormating """ & .Caption & """ '"
                    .TooltipText = barRightClickMarker & ""
                End With
                With .Controls.Add(Type:=msoControlButton, Temporary:=True)
                    .Style = msoButtonIconAndCaption
                    .FaceId = 159
                    .Caption = "Reset to Default"
                    .OnAction = "'OnFormating """ & .Caption & """ '"
                    .TooltipText = barRightClickMarker & ""
                End With

            End With
        End With

        'position = position + 1
        'With .Controls.Add(Type:=msoControlComboBox, Before:=position, Temporary:=True)
        '    .Caption = "Keyboard shortcuts"
        '    .AddItem "ON", 1
        '    .AddItem "OFF", 2
        '    .ListIndex = 1
        '    .DropDownWidth = 30
        '    .TooltipText = barRightClickMarker & ""
        'End With

        'position = position + 1
        'With .Controls.Add(Type:=msoControlButton, Before:=position, Temporary:=True)
        '    .BeginGroup = True
        '    .Style = msoButtonIconAndCaption
        '    .Caption = ""
        '    .TooltipText = barRightClickMarker & ""
        'End With

        'position = position + 1
        'With .Controls.Add(Type:=msoControlEdit, Before:=position, Temporary:=True)
        '    .TooltipText = barRightClickMarker & ""
        'End With
    End With
End Sub

Public Sub RemoveCustomRightClickBar()
    If barRightClick Is Nothing Then
        Set barRightClick = Application.CommandBars("Cell")
        If barRightClick Is Nothing Then Exit Sub
    End If

    If barRightClick.BuiltIn Then
        barRightClick.Reset
        Exit Sub
    End If

    ' Updated default bar
#If VBA6 Then
    Debug.Assert barRightClick.Name = "Cell"
#End If

    Dim ctrl As CommandBarControl
    For Each ctrl In barRightClick.Controls
        If Left(ctrl.TooltipText, Len(barRightClickMarker)) = barRightClickMarker Then
            On Error Resume Next
                ctrl.Delete
            On Error GoTo 0
        End If
    Next

End Sub


Public Sub OnRightClickQuickMode_BeforeRightClick(ByVal Target As Range, Cancel As Boolean)
    Debug.Print "wksWithEvents_BeforeRightClick"

    ' Display the custom bar or the default bar
    ' depending on what has been done in AddCustomRightClickBar
    barRightClick.ShowPopup

    Cancel = True
End Sub


' ********************************************************************************
' Method exported by the right click menu (and few others)

Function ExpandFormula(ByVal src As String) As String
    Const paddingSize = 4
    Dim nxt As Long, nxtnxt As Long, _
        nextOpenParenthesis As Long, _
        nextCloseParenthesis As Long, _
        nextComma As Long, _
        result As String, _
        padding As Long

    ' Remove NewLine and spaces in already formated strings
    src = Replace(src, vbCr, "")
    src = Replace(src, vbLf, "")
    src = Replace(src, " ", "") ' This may affect contnt of string

    While src <> ""
        nextOpenParenthesis = InStr(src, "("): nextOpenParenthesis = IIf(nextOpenParenthesis <> 0, nextOpenParenthesis, Len(src))
        nextCloseParenthesis = InStr(src, ")"): nextCloseParenthesis = IIf(nextCloseParenthesis <> 0, nextCloseParenthesis, Len(src))
        nextComma = InStr(src, ","): nextComma = IIf(nextComma <> 0, nextComma, Len(src))
        nxt = Application.WorksheetFunction.min(nextOpenParenthesis, nextCloseParenthesis, nextComma)

        If nextOpenParenthesis < nextCloseParenthesis And nextCloseParenthesis < 17 Then
            ' This is a short function call print it as a text
            ' No impact on padding, but we must coorect actual nxt
            nxt = nextCloseParenthesis
            result = result & vbCrLf & String(paddingSize * padding, " ") & Left(src, nxt)
        ElseIf nxt = nextOpenParenthesis Then
            ' Print until the opening parenthesis (inclusive), then increased padding
            result = result & vbCrLf & String(paddingSize * padding, " ") & Left(src, nxt)
            padding = padding + 1
        ElseIf nxt = nextCloseParenthesis Then
            If nxt = 1 Then
                ' Parenthesis aften another token, print at the end of the line and decrease padding
                result = result & ")"
            Else
                ' Print until the opening parenthesis (exclusive), then new line with decreased padding
                result = result & vbCrLf & String(paddingSize * padding, " ") & Left(src, nxt - 1)
                result = result & vbCrLf & String(paddingSize * (padding - 1), " ") & ")"
            End If
            padding = padding - 1
        Else
            If nxt < 7 Then
                ' Print until the comma (inclusive) at the end of the previous line
                result = result & Left(src, nxt)
            Else
                ' Print until the comma (inclusive) on a new line
                result = result & vbCrLf & String(paddingSize * padding, " ") & Left(src, nxt)
            End If
        End If
        src = Right(src, Len(src) - nxt)
    Wend

    ExpandFormula = result
End Function

' Since a name with an invalide Ref is useless we can remove them
Public Sub ListNamesWithInvalidRefs(wb As Workbook, Optional bRemoveThem As Boolean = False)
    Dim msg As String: msg = "Range names with #Ref targets." & vbNewLine _
        & "Click Ok to delete." & vbNewLine & vbNewLine
    Dim i As Long
    For i = wb.Names.Count To 1 Step -1
        Dim ref As String
        ref = wb.Names.Item(i).RefersTo
        If InStr(1, ref, "#REF") <> 0 Then
            msg = msg & wb.Names(i).Name & vbTab & wb.Names(i).RefersTo & vbNewLine
            If bRemoveThem Then wb.Names.Item(i).Delete
        End If
    Next
    If Not bRemoveThem Then
        Dim action: action = MsgBox(msg, vbOKCancel, "Range names with #Ref targets.")
        If action = vbOK Then ListNamesWithInvalidRefs wb, True
    End If
End Sub

' Since a name with Global scope is coupling sheets together we should remove them
Public Sub ListNamesWithGlobalScope(wb As Workbook, Optional bRemoveThem As Boolean = False)
    Dim msg As String: msg = "Range names with Global Scope." & vbNewLine _
        & "Click Ok to delete." & vbNewLine & vbNewLine
    Dim i As Long
    For i = wb.Names.Count To 1 Step -1
        If InStr(wb.Names.Item(i).Name, "!") = 0 Then
            msg = msg & wb.Names(i).Name & vbTab & wb.Names(i).RefersTo & vbNewLine
            If bRemoveThem Then wb.Names.Item(i).Delete
        End If
    Next
    If Not bRemoveThem Then
        Dim action: action = MsgBox(msg, vbOKCancel, "Range names with Global Scope.")
        If action = vbOK Then ListNamesWithGlobalScope wb, True
    End If
End Sub

Public Sub OnReduceFormulaArray(rng As Range)
    Dim topLeftCell As Range: Set topLeftCell = rng.Cells(1, 1)
    If Not topLeftCell.HasArray Then Exit Sub

    Dim rng_inter As Range
    Set rng_inter = Application.Intersect(rng, topLeftCell.CurrentArray)
    If rng_inter.Address <> rng.Address Then
        MsgBox "The selection is not fully contained in the Current Formula Array."
        Application.Union(rng, topLeftCell.CurrentArray).Select
        Exit Sub
    End If

    Dim formula_tmp As String
    formula_tmp = topLeftCell.CurrentArray.FormulaArray
    topLeftCell.CurrentArray.Formula = ""
    SetRangeFormulaArray rng, formula_tmp
End Sub


' Read / toggle the state True/False of TimeDisplay
Public Function OnRefreshCalcTimeDisplay(Optional bToggle As Boolean = False) As Boolean
    ' Static Variable
    Static bRefreshCalcTimeDisplay As Boolean


    Dim cmdbar  As CommandBar
    Set cmdbar = Application.CommandBars("Cell").Controls("Refresh").CommandBar
    Dim ctrl As CommandBarControl
    ' Set ctrl = cmdbar.Controls(7)
    Set ctrl = cmdbar.Controls("Display Time")              ' amperset are skipped
    '       OR
    ' Dim ctrl As CommandBarButton
    ' Set ctrl = CommandBars.ActionControl

    'bRefreshCalcTimeDisplay = (tmp.ListIndex = 1)          ' (msoControlDropdown) alternatively we can use .Text
    bRefreshCalcTimeDisplay = (ctrl.State = msoButtonDown)  ' (msoControlButton)


    If bToggle Then bRefreshCalcTimeDisplay = Not bRefreshCalcTimeDisplay
    If bRefreshCalcTimeDisplay Then
        ctrl.State = msoButtonDown          ' a check icon is added to the menu entry (equivalent msoButtonMixed)
    Else
        ctrl.State = msoButtonUp            ' check icon is removed
    End If

    OnRefreshCalcTimeDisplay = bRefreshCalcTimeDisplay
End Function

' Read the state of Activate
Public Function OnRefreshCalcActivate() As String

    Dim cmdbar  As CommandBar
    Set cmdbar = Application.CommandBars("Cell").Controls("Refresh").CommandBar
    Dim ctrl As CommandBarControl
    ' Set ctrl = cmdbar.Controls(7)
    Set ctrl = cmdbar.Controls("Activate")
    '     OR
    ' CommandBars.ActionControl is the source of the event (eg msoControlButton or msoControlDropdown)
    ' Dim ctrl As CommandBarControl
    ' Set ctrl = CommandBars.ActionControl

    OnRefreshCalcActivate = ctrl.Text            ' (msoControlDropdown)

End Function

Public Sub OnRefresh(what As String)
    Dim t1 As Long
    Dim t2 As Long
    t1 = GetTickCount()

    ' Should we disable events ?        Application.EnableEvents
    Dim bCalcDisp As Boolean: bCalcDisp = OnRefreshCalcTimeDisplay()
    Dim bActivate As Boolean: If InStr(OnRefreshCalcActivate, "Initial") = 0 Then bActivate = True
    Dim bCalcHidden As Boolean: bCalcHidden = False
    Dim shInit As Worksheet: Set shInit = ActiveSheet

    If LCase(what) = LCase("Range") Then
        RefreshRange Selection
    ElseIf LCase(what) = "openworkbooks" Then
        RefreshOpenWorkbooks bActivate, bCalcHidden
    ElseIf LCase(what) = "lefttoright" Then
        RefreshLeftToRight ActiveSheet.Parent, bActivate, bCalcHidden
    ElseIf LCase(what) = "lefttosheet" Then
        RefreshLeftToSheet ActiveSheet, bActivate, bCalcHidden
    ElseIf LCase(what) = "sheettoright" Then
        RefreshSheetToRight ActiveSheet, bActivate, bCalcHidden
    ElseIf LCase(what) = "selectedsheets" Then
        RefreshSelectedSheets ActiveSheet, bActivate
    ElseIf LCase(what) = LCase("CalcTimeDisplay") Then
        Dim tmp1: tmp1 = OnRefreshCalcTimeDisplay(True)
    ElseIf LCase(what) = LCase("CalcActivate") Then
        ' this bit is never called, OnAction has no effect for msoControlDropdown
        ' however the drop down can be read when executing a refresh command
        Dim tmp2: tmp2 = OnRefreshCalcActivate()
    Else
        Err.Raise vbObjectError, , "Error OnRefresh: didn't understand what was to refresh (" & what & ")."
    End If

    ' Initial sheet is restored (optional)
    If InStr(OnRefreshCalcActivate, "Final") = 0 Then shInit.Activate

    t2 = GetTickCount()
    If bCalcDisp Then
        MsgBox "Calculation Time: " & (t2 - t1) / 1000 & " seconds"
    End If
End Sub

Public Sub RefreshOpenWorkbooks( _
    Optional bCalcDisp As Boolean = True, _
    Optional bCalcHidden As Boolean = False _
)
    Dim wb As Workbook
	For Each wb In Workbooks
		Call RefreshLeftToRight(wb, bCalcDisp, bCalcHidden)
		If Not wb.ReadOnly Then wb.Save
	Next
    Application.StatusBar = False
End Sub

Public Sub RefreshLeftToRight(wb As Workbook, _
    Optional bCalcDisp As Boolean = True, _
    Optional bCalcHidden As Boolean = False _
)
    Dim sh As Worksheet
    For Each sh In wb.Worksheets
        If sh.Visible = xlSheetVisible Or bCalcHidden Then
            Application.StatusBar = "Refreshing " & sh.Name & "..."
            If bCalcDisp Then sh.Activate
            sh.Calculate
        End If
    Next
    Application.StatusBar = False
    'src.Activate
    Dim i As Long
    For i = 1 To 3
        Beep
    Next i
End Sub

Public Sub RefreshLeftToSheet(toSheet As Worksheet, _
    Optional bCalcDisp As Boolean = True, _
    Optional bCalcHidden As Boolean = False _
)
    Dim wb As Workbook: Set wb = toSheet.Parent
    Dim calc As Boolean: calc = True

    Dim sh As Worksheet
    For Each sh In wb.Worksheets
        If calc Then
            If sh.Visible = xlSheetVisible Or bCalcHidden Then
                Application.StatusBar = "Refreshing " & sh.Name & "..."
                If bCalcDisp Then sh.Activate
                sh.Calculate
            End If
        End If
        If sh.CodeName = toSheet.CodeName Then calc = False
    Next
    Application.StatusBar = False
    'src.Activate
    Dim i As Long
    For i = 1 To 3
        Beep
    Next i
End Sub

Public Sub RefreshSheetToRight(from As Worksheet, _
    Optional bCalcDisp As Boolean = True, _
    Optional bCalcHidden As Boolean = False _
)
    Dim wb As Workbook: Set wb = from.Parent
    Dim calc As Boolean: calc = False

    Dim sh As Worksheet
    For Each sh In wb.Worksheets
        If sh.CodeName = from.CodeName Then calc = True
        If calc Then
            If sh.Visible = xlSheetVisible Or bCalcHidden Then
                Application.StatusBar = "Refreshing " & sh.Name & "..."
                If bCalcDisp Then sh.Activate
                sh.Calculate
            End If
        End If
    Next
    Application.StatusBar = False
    'src.Activate
    Dim i As Long
    For i = 1 To 3
        Beep
    Next i
End Sub

Public Sub RefreshSelectedSheets(sh0 As Worksheet, Optional bCalcDisp As Boolean = True)
    Dim wb As Workbook: Set wb = sh0.Parent

    Dim wd As Window
    Set wd = ActiveWindow
    Set wd = wb.Windows(1)
    Dim selected_sheets As Sheets
    Set selected_sheets = wd.SelectedSheets

    Dim sh As Worksheet
    For Each sh In selected_sheets
        Application.StatusBar = "Refreshing " & sh.Name & "..."
        sh.Select
        If bCalcDisp Then sh.Activate
        sh.Calculate
    Next
    Application.StatusBar = False
    sh0.Select
    sh0.Activate
    Dim i As Long
    For i = 1 To 3
        Beep
    Next i
End Sub

Public Sub RefreshRange(rng As Range)
    On Error GoTo RefreshRangeFailed
    ' Optimistic approach
    rng.Calculate
    Exit Sub

RefreshRangeFailed:
    RefreshRangeArray rng
End Sub

Private Sub RefreshRangeArray(rng As Range)

    Set rng = MinimalCoveringRange(rng)
    rng.Select

    Dim q
    q = MsgBox("Excel can not Calculate a selection containing a partial array formula Range." _
        & vbNewLine & "The selected Range has been modified to the minimal calculable covering Range." _
        & vbNewLine & "Do you want the new selection to be Calculated ?" _
        & vbNewLine _
        & vbNewLine & "To avoid this question, select the CurrentRegion (CTRL-*) or CurrentArray (CTRL-/) first.", _
        vbQuestion + vbYesNo)

    If q = vbYes Then rng.Calculate
End Sub


Private Function MinimalCoveringRange(rng As Range) As Range
    Dim cover As Range
    Set cover = rng

    Dim c As Range
    For Each c In rng.Cells
        If c.HasArray Then
            Set cover = Union(cover, c.CurrentArray)
        End If
    Next
    Set MinimalCoveringRange = cover
End Function

Public Sub UncheckAllCheckBoxes(sh As Worksheet)
    Dim i, n: n = sh.CheckBoxes.Count
    Dim cb As CheckBox
    For i = 1 To n
        Set cb = sh.CheckBoxes(i)
        cb.Value = -4146
    Next
End Sub

Private Sub SetRangeFormulaArray(rng As Range, formula_txt As String)
    ' Excel may fail to set a big formula array directly but is happy to do a substitution
    Dim formula_tmp: formula_tmp = "=""qwertyuiopqsdfg"""
    rng.FormulaArray = formula_tmp
    rng.Replace formula_tmp, formula_txt
End Sub
