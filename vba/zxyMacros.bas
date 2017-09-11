Attribute VB_Name = "zxyMacros"
Sub getFullName()
    On Error Resume Next
    MsgBox (Application.ActiveDocument.FullName)
    Dim clipboard As MSForms.DataObject
    Set clipboard = New MSForms.DataObject
    clipboard.SetText Application.ActiveDocument.FullName
    clipboard.PutInClipboard
End Sub

Function f_��ݼ�������()
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

Function f_������������()
    Dim kbLoop As KeyBinding
    CustomizationContext = NormalTemplate
    For Each kbLoop In KeyBindings
        Selection.InsertAfter kbLoop.Command & vbTab _
        & kbLoop.KeyString & vbCr
        Selection.Collapse Direction:=wdCollapseEnd
    Next kbLoop
End Function

Function f_��ݼ�����ʾ()
    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.���ֿ��", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyF)
End Function


Function f_��ݼ�������()

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
    Command:="Normal.zxyMacros.�޸�ʽ����", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyC)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.zxyȡ���߿�", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyF10)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.s_highlight", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyF8)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.�޸�ʽճ��", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyV)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.��չ������", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKey1)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.��չ����β", _
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
    Command:="Normal.zxyMacros.���Ƹ�ʽ", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKey3)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.ճ����ʽ", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKey4)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.���Ƹ�ʽ", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyG)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.ճ����ʽ", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyG)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.�л���ɫ", _
    KeyCode:=BuildKeyCode(wdKeyF11)

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyW), KeyCategory:= _
    wdKeyCategoryCommand, Command:="DeleteBackWord"

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.ɾ����ѡ", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyD)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.�л���ʾ���ĵ��ṹͼ", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyF4)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.s_Clearallhighlight", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyF8)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.s_highlightAll", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey8)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.��һ��ע", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyPeriod)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.��һ��ע", _
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

' ����ѡ�������»�ԭΪԭ��λ��,
' ������ѡ�����ڵ�ǰ��Ļ�ɼ�
' ���������н�ֹ��Ļ����
Function f_RestoreByRange(ByRef rng As Range)
    Application.ScreenUpdating = False
    Selection.SetRange _
    Start:=rng.Start, End:=rng.End
    ActiveWindow.ScrollIntoView Selection.Range, True
    Application.ScreenUpdating = True
End Function

' ������ѡ������ı���ȥ����β�Ļس�����ո�
Function f_TrimRight() As String
    Dim Text As String          '������ѡ�����ı�
    Text = Selection.Text

    ' ȥ����β�Ļس���
    If right(Text, 1) = Chr(13) Then
        Text = Left(Text, Len(Text) - 1)
    End If

    ' ȥ����β�Ŀո�
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
    Dim sel_text As String      '������ѡ�����ı�
    Dim range_save As Range     '������ѡ����Χ
    Dim rngHighlight As Range
    Dim currentpage As Integer, pages As Integer        '��ǰҳ������
    Dim dohighlightfrom, dohighlightto As Long
    Dim bhighlight As Boolean

    Options.DefaultHighlightColorIndex = wdYellow

    ' ������ǰ��ѡ����ķ�Χ�Ա������ָ�
    Set range_save = Selection.Range

    ' �����ѡ����Ϊ�գ�����ĵ������и������˳�
    ' fixme, ��������и�������ֻ�ǵ�ǰҳ��ǰ��ҳ�еĸ��� ?
    If Selection.Range.Start = Selection.Range.End Then
        f_Clearallhighlight (True)
        f_RestoreByRange range_save
    End If

    ' ������ѡ������ı���ȥ����β�Ļس�����ո�
    sel_text = f_TrimRight()

    ' ��ǰ��ѡ�Ƿ����
    bhighlight = f_IsHighlight()

    currentpage = Selection.Information(wdActiveEndPageNumber)
    pages = Selection.Information(wdNumberOfPagesInDocument)

    dohighlightfrom = 0
    dohighlightto = ActiveDocument.Content.End

    Set rngHighlight = ActiveDocument.Range(Start:=dohighlightfrom, End:=dohighlightto)

    ' �����ǰ��ѡ���򱻸����������ǰ����
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
    Dim sel_text As String      '������ѡ�����ı�
    Dim range_save As Range     '������ѡ����Χ
    Dim rngHighlight As Range
    Dim currentpage As Integer, pages As Integer        '��ǰҳ������
    Dim backpages As Integer, forwardpages As Integer   '����ǰ��ƫ��ҳ��
    Dim dohighlightfrom, dohighlightto As Long
    Dim bhighlight As Boolean

    Options.DefaultHighlightColorIndex = wdYellow

    backpages = 2
    forwardpages = 5

    ' ������ǰ��ѡ����ķ�Χ�Ա������ָ�
    Set range_save = Selection.Range

    ' �����ѡ����Ϊ�գ�����ĵ������и������˳�
    ' fixme, ��������и�������ֻ�ǵ�ǰҳ��ǰ��ҳ�еĸ��� ?
    If Selection.Range.Start = Selection.Range.End Then
        f_Clearallhighlight (True)
        f_RestoreByRange range_save
    End If

    ' ������ѡ������ı���ȥ����β�Ļس�����ո�
    sel_text = f_TrimRight()

    ' ��ǰ��ѡ�Ƿ����
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

    ' �����ǰ��ѡ���򱻸����������ǰ����
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

' ��ԭ������������������һ��
Sub FindTextPrev()
    With Selection.Find
        .Forward = False
        .Wrap = wdFindAsk
        .Execute
    End With
End Sub

' ��ԭ������������������һ��
Sub FindTextNext()
    With Selection.Find
        .Forward = True
        .Wrap = wdFindAsk
        .Execute
    End With
End Sub

' ����ǰѡ����Ϊ�µ�����������������һ��
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

' ����ǰѡ����Ϊ�µ�����������������һ��
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

Sub ���ֿ��()
    ActiveWindow.ActivePane.View.Zoom.PageFit = wdPageFitTextFit
End Sub

Sub �޸�ʽ����()
    Dim data As New DataObject
    If Selection.Start <> Selection.End Then
        data.SetText (Selection.Text)
        data.PutInClipboard
    End If
End Sub

Sub ��չ������()
    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    'Application.Run MacroName:="Normal.zxyMacros.�޸�ʽ����"
End Sub

Sub ��չ����β()
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    ' ȥ����β�Ļس���
    If right(Selection.Text, 1) = Chr(13) Then
        Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    End If

    If right(Selection.Text, 1) = "" Then
        Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    End If

    While right(Selection.Text, 1) = Chr(32)
        ' whileѭ��ȥ����β�Ŀո�
        Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Wend

    ' ���ﾭ����Ҫ��չ����п�����ֱ�ӵ����Ѿ�д�õĺ꣬ģ�黯
    ' Application.Run MacroName:="Normal.zxyMacros.�޸�ʽ����"
End Sub

Sub �л�������ʾ�޶�()
    With ActiveWindow.View
        .ShowRevisionsAndComments = Not (.ShowRevisionsAndComments)
        .RevisionsView = wdRevisionsViewFinal
    End With
End Sub

Sub �޸�ʽճ��()
    On Error Resume Next
doit:
    Selection.PasteAndFormat (wdFormatPlainText)
    On Error GoTo doit
End Sub

Sub ���Ƹ�ʽ()
    Selection.CopyFormat
End Sub

Sub ճ����ʽ()
    Selection.PasteFormat
End Sub

Sub ɾ����ѡ()
    Selection.Delete Unit:=wdCharacter, Count:=1
End Sub

Sub �л���ɫ()
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

Sub �л���ʾ���ĵ��ṹͼ()
    On Error Resume Next
    ActiveWindow.DocumentMap = Not (ActiveWindow.DocumentMap)
    On Error Resume Next
End Sub

Sub s_Add��ݼ�()
    f_��ݼ�������
    f_��ݼ�����ʾ
    f_��ݼ�������
End Sub

Sub AutoExec()
    Application.DisplayAlerts = wdAlertsNone

    'CustomizationContext = NormalTemplate
    s_Add��ݼ�
    Application.CommandBars("Research").Enabled = False
    On Error Resume Next
    ActiveWindow.View.DisplayPageBoundaries = False

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.��һ����", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyShift, wdKeyTab)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyMacros.��һ����", _
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

Sub zxyȡ���߿�()
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

Sub ��һ��ע()
    With Application.Browser
        .Target = wdBrowseComment
        .Next
    End With
End Sub

Sub ��һ��ע()
    With Application.Browser
        .Target = wdBrowseComment
        .Previous
    End With
End Sub

Sub ��һ����()
    i = ActiveWindow.Index
    If i = Windows.Count Then
        i = 1
    Else
        i = i + 1
    End If
    Windows(i).Activate
End Sub

Sub ��һ����()
    i = ActiveWindow.Index - 1
    If i = 0 Then
        i = Windows.Count
    End If
    Windows(i).Activate
End Sub

