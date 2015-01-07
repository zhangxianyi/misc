Attribute VB_Name = "NewMacros"
Function f_快捷键·查找()
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.FindTextPrev", _
    KeyCode:=BuildKeyCode(wdKeyF3)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.FindTextNext", _
    KeyCode:=BuildKeyCode(wdKeyF4)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.FindSelTextPrev", _
    KeyCode:=BuildKeyCode(wdKeyShift, wdKeyF3)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.FindSelTextNext", _
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
    Command:="Normal.NewMacros.文字宽度", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyShift, wdKeyF)
End Function

Function f_快捷键·其它()

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.无格式拷贝", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyC)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.zxy取消边框", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyF10)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.s_highlight", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyF8)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.无格式粘贴", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyV)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.扩展至行首", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKey1)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.扩展至行尾", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKey2)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="正文", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyShift, wdKeyN)

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyZ, wdKeyAlt), KeyCategory:= _
    wdKeyCategoryCommand, Command:="EditUndo"

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyH, wdKeyAlt), KeyCategory:= _
    wdKeyCategoryCommand, Command:="WordLeft"

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyL, wdKeyAlt), KeyCategory:= _
    wdKeyCategoryCommand, Command:="WordRight"

    'KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyH, wdKeyControl), KeyCategory:= _
    'wdKeyCategoryCommand, Command:="CharLeft"
    
    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyH, wdKeyControl), KeyCategory:= _
    wdKeyCategoryMacro, Command:="Normal.NewMacros.zxyViewMode"

    'KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyL, wdKeyControl), KeyCategory:= _
    'wdKeyCategoryCommand, Command:="CharRight"

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyA, wdKeyAlt), KeyCategory:= _
    wdKeyCategoryCommand, Command:="StartOfLine"

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyE, wdKeyAlt), KeyCategory:= _
    wdKeyCategoryCommand, Command:="EndOfLine"

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyX, wdKeyAlt), KeyCategory:= _
    wdKeyCategoryCommand, Command:="EditCut"

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.复制格式", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKey3)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.粘贴格式", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKey4)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.复制格式", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyShift, wdKeyG)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.粘贴格式", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyG)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.切换颜色", _
    KeyCode:=BuildKeyCode(wdKeyF11)

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyW, wdKeyAlt), KeyCategory:= _
    wdKeyCategoryCommand, Command:="DeleteBackWord"

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.删除所选", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyD)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="menu1", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyShift, wdKey1)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="menu2", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyShift, wdKey2)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="代码样式-六号", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKey8)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.切换显示·文档结构图", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyF4)
End Function


Function f_Clearallhighlight()
    With ActiveDocument.Range(Start:=0, End:=0).Find
    .ClearFormatting
    .Highlight = True
    With .Replacement
        .ClearFormatting
        .Highlight = False
    End With

    .Execute Replace:=wdReplaceAll, _
    Format:=True, _
    MatchWholeWord:=True
    End With
End Function


' 将所选区域重新还原为原来位置,
' 并将所选区域在当前屏幕可见
' 操作过程中禁止屏幕更新
Function f_RestoreByRange(ByRef rng As Range)
    Application.ScreenUpdating = False
    Selection.SetRange _
    Start:=rng.Start, _
    End:=rng.End

    ActiveWindow.ScrollIntoView Selection.Range, True
    Application.ScreenUpdating = True
End Function

' 保留所选区域的文本并去除行尾的回车符与空格
Function f_TrimRight() As String
    Dim text As String          '保存所选区域文本
    text = Selection.text

    ' 去除行尾的回车符
    If Right(text, 1) = Chr(13) Then
    text = Left(text, Len(text) - 1)
    End If

    ' 去除行尾的空格
    While Right(text, 1) = Chr(32)
    text = Left(text, Len(text) - 1)
    Wend

    f_TrimRight = text
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

Sub s_highlight()
    Dim sel_text As String          '保存所选区域文本
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
    f_Clearallhighlight
    f_RestoreByRange range_save
    Exit Function
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
    dohighlightfrom = Selection.GoTo(what:=wdGoToPage, _
    which:=wdGoToPrevious, _
    Count:=backpages).Start
    End If

    If (currentpage + forwardpages + 1 >= pages) Then
    dohighlightto = ActiveDocument.Content.End
    Else
    dohighlightto = Selection.GoTo(what:=wdGoToPage, _
    which:=wdGoToNext, _
    Count:=forwardpages + 1).End
    End If

    Set rngHighlight = ActiveDocument.Range(Start:=dohighlightfrom, _
    End:=dohighlightto)

    ' 如果当前所选区域被高亮，清除当前高亮
    With rngHighlight.Find
    .ClearFormatting
    '.Highlight = True
    .text = sel_text
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
    .Execute
    End With
End Sub

' 按原来的搜索条件搜索下一个
Sub FindTextNext()
    With Selection.Find
    .Forward = True
    .Execute
    End With
End Sub

' 将当前选择作为新的搜索条件，搜索上一个
Sub FindSelTextPrev()
    If Selection.Start = Selection.End Then
    FindTextPrev
    End If
    sel_text = f_TrimRight()
    If Selection.Start = Selection.End Then
    FindTextPrev
    End If
    With Selection.Find
    .Forward = False
    .ClearFormatting
    .MatchWholeWord = False
    .MatchCase = False
    .Wrap = wdFindAsk
    .Execute
    End With
End Sub

' 将当前选择作为新的搜索条件，搜索下一个
Sub FindSelTextNext()
    If Selection.Start = Selection.End Then
    FindTextNext
    End If
    sel_text = f_TrimRight()
    If Selection.Start = Selection.End Then
    FindTextNext
    End If
    With Selection.Find
    .Forward = True
    .ClearFormatting
    .MatchWholeWord = False
    .MatchCase = False
    .Wrap = wdFindAsk
    .Execute
    End With
End Sub

Sub 文字宽度()
    ActiveWindow.ActivePane.View.Zoom.PageFit = wdPageFitTextFit
End Sub

Sub 无格式拷贝()
    Dim data As New DataObject
    If Selection.Start <> Selection.End Then
    data.settext Selection.text
    data.putinclipboard
    End If
End Sub

Sub 扩展至行首()
    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    'Application.Run MacroName:="Normal.NewMacros.无格式拷贝"
End Sub

Sub 扩展至行尾()
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    ' 去除行尾的回车符
    If Right(Selection.text, 1) = Chr(13) Then
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    End If

    While Right(Selection.text, 1) = Chr(32)
    ' while循环去除行尾的空格
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Wend

    ' 这里经常是要扩展后进行拷贝，直接调用已经写好的宏，模块化
    ' Application.Run MacroName:="Normal.NewMacros.无格式拷贝"
End Sub

Sub 切换隐藏显示修订()
    With ActiveWindow.View
    .ShowRevisionsAndComments = Not (.ShowRevisionsAndComments)
    .RevisionsView = wdRevisionsViewFinal
    End With
End Sub

Sub 无格式粘贴()
    Selection.PasteAndFormat (wdFormatPlainText)
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
    ElseIf Selection.Font.Color = wdColorBlue Then
    Selection.Font.Color = wdColorBlack
    End If
End Sub

Sub 切换显示·文档结构图()
    ActiveWindow.DocumentMap = Not (ActiveWindow.DocumentMap)
End Sub

Sub s_Add快捷键()
    f_快捷键·查找
    f_快捷键·显示
    f_快捷键·其它
End Sub

Sub AutoExec()
    CustomizationContext = NormalTemplate
    s_Add快捷键
End Sub

Sub AutoOpen()
    'ActiveWindow.View.ShowSpaces = True
    ActiveWindow.View.DisplayPageBoundaries = False
End Sub

Sub AutoNew()
    'ActiveWindow.View.ShowSpaces = True
    ActiveWindow.View.DisplayPageBoundaries = False
End Sub

Sub zxyViewMode()
    ActiveWindow.View.DisplayPageBoundaries = False
End Sub


Sub 宏1()
Attribute 宏1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.宏1"
'
' 宏1 宏
'
'
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
