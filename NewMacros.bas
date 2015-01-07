Attribute VB_Name = "NewMacros"
Function f_��ݼ�������()
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
    Command:="Normal.NewMacros.���ֿ��", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyShift, wdKeyF)
End Function

Function f_��ݼ�������()

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.�޸�ʽ����", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyC)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.zxyȡ���߿�", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyF10)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.s_highlight", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyF8)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.�޸�ʽճ��", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyV)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.��չ������", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKey1)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.��չ����β", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKey2)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="����", _
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
    Command:="Normal.NewMacros.���Ƹ�ʽ", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKey3)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.ճ����ʽ", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKey4)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.���Ƹ�ʽ", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyShift, wdKeyG)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.ճ����ʽ", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyG)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.�л���ɫ", _
    KeyCode:=BuildKeyCode(wdKeyF11)

    KeyBindings.Add KeyCode:=BuildKeyCode(wdKeyW, wdKeyAlt), KeyCategory:= _
    wdKeyCategoryCommand, Command:="DeleteBackWord"

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.ɾ����ѡ", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyD)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="menu1", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyShift, wdKey1)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="menu2", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKeyShift, wdKey2)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="������ʽ-����", _
    KeyCode:=BuildKeyCode(wdKeyAlt, wdKey8)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.NewMacros.�л���ʾ���ĵ��ṹͼ", _
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


' ����ѡ�������»�ԭΪԭ��λ��,
' ������ѡ�����ڵ�ǰ��Ļ�ɼ�
' ���������н�ֹ��Ļ����
Function f_RestoreByRange(ByRef rng As Range)
    Application.ScreenUpdating = False
    Selection.SetRange _
    Start:=rng.Start, _
    End:=rng.End

    ActiveWindow.ScrollIntoView Selection.Range, True
    Application.ScreenUpdating = True
End Function

' ������ѡ������ı���ȥ����β�Ļس�����ո�
Function f_TrimRight() As String
    Dim text As String          '������ѡ�����ı�
    text = Selection.text

    ' ȥ����β�Ļس���
    If Right(text, 1) = Chr(13) Then
    text = Left(text, Len(text) - 1)
    End If

    ' ȥ����β�Ŀո�
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
    Dim sel_text As String          '������ѡ�����ı�
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
    f_Clearallhighlight
    f_RestoreByRange range_save
    Exit Function
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

    ' �����ǰ��ѡ���򱻸����������ǰ����
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

' ��ԭ������������������һ��
Sub FindTextPrev()
    With Selection.Find
    .Forward = False
    .Execute
    End With
End Sub

' ��ԭ������������������һ��
Sub FindTextNext()
    With Selection.Find
    .Forward = True
    .Execute
    End With
End Sub

' ����ǰѡ����Ϊ�µ�����������������һ��
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

' ����ǰѡ����Ϊ�µ�����������������һ��
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

Sub ���ֿ��()
    ActiveWindow.ActivePane.View.Zoom.PageFit = wdPageFitTextFit
End Sub

Sub �޸�ʽ����()
    Dim data As New DataObject
    If Selection.Start <> Selection.End Then
    data.settext Selection.text
    data.putinclipboard
    End If
End Sub

Sub ��չ������()
    Selection.HomeKey Unit:=wdLine, Extend:=wdExtend
    'Application.Run MacroName:="Normal.NewMacros.�޸�ʽ����"
End Sub

Sub ��չ����β()
    Selection.EndKey Unit:=wdLine, Extend:=wdExtend
    ' ȥ����β�Ļس���
    If Right(Selection.text, 1) = Chr(13) Then
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    End If

    While Right(Selection.text, 1) = Chr(32)
    ' whileѭ��ȥ����β�Ŀո�
    Selection.MoveLeft Unit:=wdCharacter, Count:=1, Extend:=wdExtend
    Wend

    ' ���ﾭ����Ҫ��չ����п�����ֱ�ӵ����Ѿ�д�õĺ꣬ģ�黯
    ' Application.Run MacroName:="Normal.NewMacros.�޸�ʽ����"
End Sub

Sub �л�������ʾ�޶�()
    With ActiveWindow.View
    .ShowRevisionsAndComments = Not (.ShowRevisionsAndComments)
    .RevisionsView = wdRevisionsViewFinal
    End With
End Sub

Sub �޸�ʽճ��()
    Selection.PasteAndFormat (wdFormatPlainText)
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
    ElseIf Selection.Font.Color = wdColorBlue Then
    Selection.Font.Color = wdColorBlack
    End If
End Sub

Sub �л���ʾ���ĵ��ṹͼ()
    ActiveWindow.DocumentMap = Not (ActiveWindow.DocumentMap)
End Sub

Sub s_Add��ݼ�()
    f_��ݼ�������
    f_��ݼ�����ʾ
    f_��ݼ�������
End Sub

Sub AutoExec()
    CustomizationContext = NormalTemplate
    s_Add��ݼ�
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


Sub ��1()
Attribute ��1.VB_ProcData.VB_Invoke_Func = "Normal.NewMacros.��1"
'
' ��1 ��
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
