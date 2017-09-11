Attribute VB_Name = "zxyMacros"
Sub getFullName()
    On Error Resume Next
    MsgBox (Application.ActiveDocument.FullName)
    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject
    clipboard.SetText Application.ActiveDocument.FullName
    clipboard.PutInClipboard
End Sub

Function f_快捷键·查找()
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.FindTextPrev", _
    KeyCode:=BuildKeyCode(wdKeyF3)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.FindTextNext", _
    KeyCode:=BuildKeyCode(wdKeyF4)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.FindSelTextPrev", _
    KeyCode:=BuildKeyCode(wdKeyShift, wdKeyF3)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.FindSelTextNext", _
    KeyCode:=BuildKeyCode(wdKeyShift, wdKeyF4)
End Function

Function f_输出所有命令绑定()
    Dim kbLoop As KeyBinding
    CustomizationContext = NormalTemplate
    For Each kbLoop In KeyBindings
        Selection.InsertAfter kbLoop.Command & vbTab _
        & kbLoop.KeyString & vbCr
        Selection.Collapse Direction:=wdCollapseEnd
    Next kbLoop
End Function

Function f_快捷键·显示()
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.文字宽度", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyF)
End Function


Function f_快捷键·其它()

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.zxySelectionFieldsUpdate", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyF5)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.zxyBrowseHeadingNext", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyPageDown)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.zxyBrowseHeadingPrevious", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyPageUp)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.switchOutlineView", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyF12)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.getFullName", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyF9)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.getFullName", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyF9)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.无格式拷贝", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyC)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.zxy取消边框", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyF10)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.s_highlight", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyF8)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.无格式粘贴", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyV)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.扩展至行首", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKey1)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.扩展至行尾", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKey2)

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyZ), KeyCategory:= _
    wdKeyCategoryCommand, Command:="EditUndo"

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyH), KeyCategory:= _
    wdKeyCategoryCommand, Command:="WordLeft"

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyL), KeyCategory:= _
    wdKeyCategoryCommand, Command:="WordRight"

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyH), KeyCategory:= _
    wdKeyCategoryCommand, Command:="EditReplace"

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyH), KeyCategory:= _
    wdKeyCategoryMacro, Command:="Normal.zxyMacros.zxyViewMode"

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyA), KeyCategory:= _
    wdKeyCategoryCommand, Command:="StartOfLine"

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyE), KeyCategory:= _
    wdKeyCategoryCommand, Command:="EndOfLine"

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyX), KeyCategory:= _
    wdKeyCategoryCommand, Command:="EditCut"

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.复制格式", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKey3)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.粘贴格式", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKey4)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.复制格式", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyG)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.粘贴格式", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyG)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.切换颜色", _
    KeyCode:=BuildKeyCode(wdKeyF11)

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyW), KeyCategory:= _
    wdKeyCategoryCommand, Command:="DeleteBackWord"

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.删除所选", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyD)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.切换显示·文档结构图", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyF4)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.s_Clearallhighlight", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyF8)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.s_highlightAll", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey8)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.下一批注", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyPeriod)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.上一批注", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyComma)

End Function

Function f_Clearallhighlight(bwhole)
    With ActiveDocument.Range(Start:=0, End:=0).Find
        .ClearFormatting
        .Highlight = True
        With .Replacement
            .ClearFormatting
            .Highlight = False
        End With

        .Execute Replace:=wdReplaceAll, _
        Format:=True, _
        MatchWholeWord:=bwhole
    End With
End Function

Sub s_Clearallhighlight()
    Call f_Clearallhighlight(False)
End Sub

' 将所选区域重新还原为原来位置,
' 并将所选区域在当前屏幕可见
' 操作过程中禁止屏幕更新
Function f_RestoreByRange(ByRef rng As Range)
    Application.ScreenUpdating = False
    Selection.SetRange _
    Start:=rng.Start, End:=rng.End
    ActiveWindow.ScrollIntoView Selection.Range, True
    Application.ScreenUpdating = True
End Function

' 保留所选区域的文本并去除行尾的回车符与空格
Function f_TrimRight() As String
    Dim Text As String          '保存所选区域文本
    Text = Selection.Text

    ' 去除行尾的回车符
    If right(Text, 1) = Chr(13) Then
        Text = Left(Text, Len(Text) - 1)
    End If

    ' 去除行尾的空格
    While right(Text, 1) = Chr(32)
        Text = Left(Text, Len(Text) - 1)
    Wend

    f_TrimRight = Text
End Function

Function f_IsHighlight() As Boolean
    colorIdx = Selection.Range.HighlightColorIndex
    If colorIdx = wdNoHighlight Or _
        colorIdx = 9999999 Then
        f_IsHighlight = False
    Else
        f_IsHighlight = True
    End If
End Function

Sub s_highlightAll()
    Dim sel_text As String      '保存所选区域文本
    Dim range_save As Range     '保存所选区域范围
    Dim rngHighlight As Range
    Dim currentpage As Integer, pages As Integer        '当前页面总量
    Dim dohighlightfrom, dohighlightto As Long
    Dim bhighlight As Boolean

    Options.DefaultHighlightColorIndex = wdYellow

    ' 保留当前所选区域的范围以便操作后恢复
    Set range_save = Selection.Range

    ' 如果所选区域为空，清除文档中所有高亮后退出
    ' fixme, 是清除所有高亮还是只是当前页面前后几页中的高亮 ?
    If Selection.Range.Start = Selection.Range.End Then
        f_Clearallhighlight (True)
        f_RestoreByRange range_save
    End If

    ' 保留所选区域的文本并去除行尾的回车符与空格
    sel_text = f_TrimRight()

    ' 当前所选是否高亮
    bhighlight = f_IsHighlight()

    currentpage = Selection.Information(wdActiveEndPageNumber)
    pages = Selection.Information(wdNumberOfPagesInDocument)

    dohighlightfrom = 0
    dohighlightto = ActiveDocument.Content.End

    Set rngHighlight = ActiveDocument.Range(Start:=dohighlightfrom, End:=dohighlightto)

    ' 如果当前所选区域被高亮，清除当前高亮
    With rngHighlight.Find
        .ClearFormatting
        .Text = sel_text
        With .Replacement
            .ClearFormatting
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll, _
        Format:=True, _
        MatchWholeWord:=False
    End With
    f_RestoreByRange range_save
End Sub

Sub s_highlight()
    Dim sel_text As String      '保存所选区域文本
    Dim range_save As Range     '保存所选区域范围
    Dim rngHighlight As Range
    Dim currentpage As Integer, pages As Integer        '当前页面总量
    Dim backpages As Integer, forwardpages As Integer   '高亮前后偏移页数
    Dim dohighlightfrom, dohighlightto As Long
    Dim bhighlight As Boolean

    Options.DefaultHighlightColorIndex = wdYellow

    backpages = 2
    forwardpages = 5

    ' 保留当前所选区域的范围以便操作后恢复
    Set range_save = Selection.Range

    ' 如果所选区域为空，清除文档中所有高亮后退出
    ' fixme, 是清除所有高亮还是只是当前页面前后几页中的高亮 ?
    If Selection.Range.Start = Selection.Range.End Then
        f_Clearallhighlight (True)
        f_RestoreByRange range_save
    End If

    ' 保留所选区域的文本并去除行尾的回车符与空格
    sel_text = f_TrimRight()

    ' 当前所选是否高亮
    bhighlight = f_IsHighlight()

    currentpage = Selection.Information(wdActiveEndPageNumber)
    pages = Selection.Information(wdNumberOfPagesInDocument)

    If currentpage < backpages + 1 Then
        dohighlightfrom = 0
    Else
        dohighlightfrom = Selection.GoTo(What:=wdGoToPage, _
        Which:=wdGoToPrevious, _
        Count:=backpages).Start
    End If

    If (currentpage + forwardpages + 1 >= pages) Then
        dohighlightto = ActiveDocument.Content.End
    Else
        dohighlightto = Selection.GoTo(What:=wdGoToPage, _
        Which:=wdGoToNext, _
        Count:=forwardpages + 1).End
    End If

    Set rngHighlight = ActiveDocument.Range(Start:=dohighlightfrom, End:=dohighlightto)

    ' 如果当前所选区域被高亮，清除当前高亮
    With rngHighlight.Find
        .ClearFormatting
        '.Highlight = True
        .Text = sel_text
        With .Replacement
            .ClearFormatting
            .Highlight = True
        End With
        .Execute Replace:=wdReplaceAll, _
        Format:=True, _
        MatchWholeWord:=False
    End With
    f_RestoreByRange range_save
End Sub

' 按原来的搜索条件搜索上一个
Sub FindTextPrev()
    With Selection.Find
        .Forward = False
        .Wrap = wdFindAsk
        .Execute
    End With
End Sub

' 按原来的搜索条件搜索下一个
Sub FindTextNext()
    With Selection.Find
        .Forward = True
        .Wrap = wdFindAsk
        .Execute
    End With
End Sub

' 将当前选择作为新的搜索条件，搜索上一个
Sub FindSelTextPrev()
    If Selection.Start = Selection.End Then
        FindTextPrev
    End If
    'sel_text = f_TrimRight()
    sel_text = Selection.Text
    If Selection.Start = Selection.End Then
        FindTextPrev
    End If
    With Selection.Find
        .Forward = False
        .ClearFormatting
        .MatchWholeWord = False
        .MatchCase = False
        .Wrap = wdFindAsk
        .Execute FindText:=sel_text
    End With
End Sub

' 将当前选择作为新的搜索条件，搜索下一个
Sub FindSelTextNext()
    If Selection.Start = Selection.End Then
        FindTextNext
    End If
    'sel_text = f_TrimRight()
    sel_text = Selection.Text

    If Selection.Start = Selection.End Then
        FindTextNext
    End If
    With Selection.Find
        .Forward = True
        .ClearFormatting
        .MatchWholeWord = False
        .MatchCase = False
        .Wrap = wdFindAsk
        .Execute FindText:=sel_text
    End With
End Sub

Sub 文字宽度()
    ActiveWindow.ActivePane.View.Zoom.PageFit = wdPageFitTextFit
End Sub

Sub 无格式拷贝()
    Dim data As New DataObject
    If Selection.Start <> Selection.End Then
        data.SetText (Selection.Text)
        data.PutInClipboard
    End If
End Sub

Sub 扩展至行首()
    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    'Application.Run MacroName:="Normal.zxyMacros.无格式拷贝"
End Sub

Sub 扩展至行尾()
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    ' 去除行尾的回车符
    If right(Selection.Text, 1) = Chr(13) Then
        Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    End If

    If right(Selection.Text, 1) = "" Then
        Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    End If

    While right(Selection.Text, 1) = Chr(32)
        ' while循环去除行尾的空格
        Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Wend

    ' 这里经常是要扩展后进行拷贝，直接调用已经写好的宏，模块化
    ' Application.Run MacroName:="Normal.zxyMacros.无格式拷贝"
End Sub

Sub 切换隐藏显示修订()
    With ActiveWindow.View
        .ShowRevisionsAndComments = Not (.ShowRevisionsAndComments)
        .RevisionsView = wdRevisionsViewFinal
    End With
End Sub

Sub 无格式粘贴()
    On Error Resume Next
doit:
    Selection.PasteAndFormat (wdFormatPlainText)
    On Error GoTo doit
End Sub

Sub 复制格式()
    Selection.CopyFormat
End Sub

Sub 粘贴格式()
    Selection.PasteFormat
End Sub

Sub 删除所选()
    Selection.Delete Unit:=wdCharacter, Count:=1
End Sub

Sub 切换颜色()
    If Selection.Font.Color = wdColorBlack Or Selection.Font.Color = -16777216 Then
        Selection.Font.Color = wdColorRed
    ElseIf Selection.Font.Color = wdColorRed Then
        Selection.Font.Color = wdColorBlue
        '    Selection.Font.Color = wdColorDarkRed
        'ElseIf Selection.Font.Color = wdColorDarkRed Then
        '   Selection.Font.Color = wdColorBlue
    ElseIf Selection.Font.Color = wdColorBlue Then
        Selection.Font.Color = wdColorBlack
    End If
End Sub

Sub 切换显示·文档结构图()
    On Error Resume Next
    ActiveWindow.DocumentMap = Not (ActiveWindow.DocumentMap)
    On Error Resume Next
End Sub

Sub s_Add快捷键()
    f_快捷键·查找
    f_快捷键·显示
    f_快捷键·其它
End Sub

Sub AutoExec()
    Application.DisplayAlerts = wdAlertsNone

    'CustomizationContext = NormalTemplate
    s_Add快捷键
    Application.CommandBars("Research").Enabled = False
    On Error Resume Next
    ActiveWindow.View.DisplayPageBoundaries = False

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.上一窗口", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyTab)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.下一窗口", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyTab)
    Normal.zxyStyle.StyleKeybind
End Sub

Sub AutoOpen()
    Application.DisplayAlerts = wdAlertsNone
    ActiveWindow.View.ShowSpaces = False
    ActiveWindow.View.DisplayPageBoundaries = False
End Sub

Sub AutoNew()
    ActiveWindow.View.ShowSpaces = True
    ActiveWindow.View.DisplayPageBoundaries = False
End Sub

Sub zxyViewMode()
    ActiveWindow.View.DisplayPageBoundaries = Not ActiveWindow.View.DisplayPageBoundaries
End Sub

Sub zxy取消边框()
    With Selection.Tables(1)
        .Borders(wdBorderLeft).LineStyle = wdLineStyleNone
        .Borders(wdBorderRight).LineStyle = wdLineStyleNone
        .Borders(wdBorderTop).LineStyle = wdLineStyleNone
        .Borders(wdBorderBottom).LineStyle = wdLineStyleNone
        .Borders(wdBorderHorizontal).LineStyle = wdLineStyleNone
        .Borders(wdBorderVertical).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalDown).LineStyle = wdLineStyleNone
        .Borders(wdBorderDiagonalUp).LineStyle = wdLineStyleNone
        .Borders.Shadow = False
    End With
    With Options
        .DefaultBorderLineStyle = wdLineStyleSingle
        .DefaultBorderLineWidth = wdLineWidth050pt
        .DefaultBorderColor = -587137025
    End With
End Sub
Sub switchOutlineView()
    If ActiveWindow.ActivePane.View.Type = wdOutlineView Then
        ActiveWindow.ActivePane.View.Type = wdPrintView
    Else
        ActiveWindow.ActivePane.View.Type = wdOutlineView
    End If
End Sub
Sub zxyBrowseHeadingPrevious()
    With Application.Browser
        .Target = wdBrowseHeading
        .Previous
    End With
End Sub

Sub zxyBrowseHeadingNext()
    With Application.Browser
        .Target = wdBrowseHeading
        .Next
    End With
    'ActiveDocument.ActiveWindow.View.ShowHeading (2)
End Sub

Sub zxySelectionFieldsUpdate()
    Selection.Fields.Update
End Sub

Sub 下一批注()
    With Application.Browser
        .Target = wdBrowseComment
        .Next
    End With
End Sub

Sub 上一批注()
    With Application.Browser
        .Target = wdBrowseComment
        .Previous
    End With
End Sub

Sub 下一窗口()
    i = ActiveWindow.Index
    If i = Windows.Count Then
        i = 1
    Else
        i = i + 1
    End If
    Windows(i).Activate
End Sub

Sub 上一窗口()
    i = ActiveWindow.Index - 1
    If i = 0 Then
        i = Windows.Count
    End If
    Windows(i).Activate
End Sub

