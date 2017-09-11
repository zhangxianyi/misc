Attribute VB_Name = "zxyStyle"
Sub StyleKeybind()
    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="����", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyD)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="����-��������", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyZ)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="����", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyN)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="MP����1", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyNumeric1)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="MP����2", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyNumeric2)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="MP����3", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyNumeric3)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="MP����4", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyNumeric4)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="MP����5", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyNumeric5)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="H3C ���� 1", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyF1)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="H3C ���� 2", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyF2)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="H3C ���� 3", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyF3)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="���� 1", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey1)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="���� 2", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey2)
    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="���� 3", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey3)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="���� 4", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey4)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="���� 5", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey5)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="���� 6", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey6)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="���� 7", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey7)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="���� 8", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey8)

    KeyBindings.Add KeyCategory:=wdKeyCategoryStyle, _
    Command:="���� 9", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKey9)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyStyle.MP����", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyF12)

    KeyBindings.Add KeyCategory:=wdKeyCategoryMacro, _
    Command:="Normal.zxyStyle.H3C����", _
    KeyCode:=BuildKeyCode(wdKeyControl, wdKeyAlt, wdKeyF9)
End Sub

Sub MP����()
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = "%1"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0.76)
        .TabPosition = CentimetersToPoints(0.76)
        .ResetOnHigher = 0
        .StartAt = 1
        With .Font
            .Bold = True
            .Italic = False
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = 13395456
            .Size = 42
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = "Book Antiqua"
        End With
        .LinkedStyle = "MP����1"
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(2)
        .NumberFormat = "%1.%2"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.02)
        .TabPosition = CentimetersToPoints(1.02)
        .ResetOnHigher = 1
        .StartAt = 1
        With .Font
            .Bold = True
            .Italic = False
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = 13395456
            .Size = 14
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = "Book Antiqua"
        End With
        .LinkedStyle = "MP����2"
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(3)
        .NumberFormat = "%1.%2.%3"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.27)
        .TabPosition = CentimetersToPoints(1.27)
        .ResetOnHigher = 2
        .StartAt = 1
        With .Font
            .Bold = True
            .Italic = False
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = 13395456
            .Size = 12
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = "Book Antiqua"
        End With
        .LinkedStyle = "MP����3"
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(4)
        .NumberFormat = "%1.%2.%3.%4"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.27)
        .TabPosition = CentimetersToPoints(1.27)
        .ResetOnHigher = 3
        .StartAt = 1
        With .Font
            .Bold = True
            .Italic = False
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = 13395456
            .Size = 10
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = "����"
        End With
        .LinkedStyle = "MP����4"
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(5)
        .NumberFormat = "%1.%2.%3.%4.%5"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.27)
        .TabPosition = CentimetersToPoints(1.27)
        .ResetOnHigher = 4
        .StartAt = 1
        With .Font
            .Bold = True
            .Italic = False
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = 13395456
            .Size = 10
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = "����"
        End With
        .LinkedStyle = "MP����5"
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(6)
        .NumberFormat = "%1.%2.%3.%4.%5.%6"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.27)
        .TabPosition = CentimetersToPoints(1.27)
        .ResetOnHigher = 5
        .StartAt = 1
        With .Font
            .Bold = True
            .Italic = False
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = 13395456
            .Size = 10
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = "����"
        End With
        .LinkedStyle = ""
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(7)
        .NumberFormat = "%1.%2.%3.%4.%5.%6.%7"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.27)
        .TabPosition = CentimetersToPoints(1.27)
        .ResetOnHigher = 0
        .StartAt = 1
        With .Font
            .Bold = True
            .Italic = False
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = 13395456
            .Size = 10
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = "����"
        End With
        .LinkedStyle = ""
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(8)
        .NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.27)
        .TabPosition = CentimetersToPoints(1.27)
        .ResetOnHigher = 0
        .StartAt = 1
        With .Font
            .Bold = True
            .Italic = False
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = 13395456
            .Size = 10
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = "����"
        End With
        .LinkedStyle = ""
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(9)
        .NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8.%9"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleArabic
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.27)
        .TabPosition = CentimetersToPoints(1.27)
        .ResetOnHigher = 8
        .StartAt = 1
        With .Font
            .Bold = True
            .Italic = False
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = 13395456
            .Size = 10
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = "����"
        End With
        .LinkedStyle = ""
    End With
    ListGalleries(wdOutlineNumberGallery).ListTemplates(1).Name = ""
    Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
    ListGalleries(wdOutlineNumberGallery).ListTemplates(1), _
    ContinuePreviousList:=True, ApplyTo:=wdListApplyToWholeList, _
    DefaultListBehavior:=wdWord10ListBehavior
End Sub

Sub H3C����()
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(1)
        .NumberFormat = "%1."
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleLegal
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(0.76)
        .TabPosition = CentimetersToPoints(0.76)
        .ResetOnHigher = 0
        .StartAt = 1
        With .Font
            .Bold = True
            .Italic = False
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdColorDarkRed
            .Size = 15
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = "����"
        End With
        .LinkedStyle = "H3C ���� 1"
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(2)
        .NumberFormat = "%1.%2"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleLegal
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.02)
        .TabPosition = CentimetersToPoints(1.02)
        .ResetOnHigher = 1
        .StartAt = 1
        With .Font
            .Bold = True
            .Italic = False
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdColorDarkRed
            .Size = 14
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = "����"
        End With
        .LinkedStyle = "H3C ���� 2"
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(3)
        .NumberFormat = "%1.%2.%3"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleLegal
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.27)
        .TabPosition = CentimetersToPoints(1.27)
        .ResetOnHigher = 2
        .StartAt = 1
        With .Font
            .Bold = True
            .Italic = False
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdColorDarkRed
            .Size = 12
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = "����"
        End With
        .LinkedStyle = "H3C ���� 3"
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(4)
        .NumberFormat = "%1.%2.%3.%4"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleLegal
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.27)
        .TabPosition = CentimetersToPoints(1.27)
        .ResetOnHigher = 3
        .StartAt = 1
        With .Font
            .Bold = True
            .Italic = False
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdColorDarkRed
            .Size = 10
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = "����"
        End With
        .LinkedStyle = ""
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(5)
        .NumberFormat = "%1.%2.%3.%4.%5"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleLegal
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.27)
        .TabPosition = CentimetersToPoints(1.27)
        .ResetOnHigher = 4
        .StartAt = 1
        With .Font
            .Bold = True
            .Italic = False
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdColorDarkRed
            .Size = 10
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = "����"
        End With
        .LinkedStyle = ""
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(6)
        .NumberFormat = "%1.%2.%3.%4.%5.%6"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleLegal
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.27)
        .TabPosition = CentimetersToPoints(1.27)
        .ResetOnHigher = 5
        .StartAt = 1
        With .Font
            .Bold = True
            .Italic = False
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdColorDarkRed
            .Size = 10
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = "����"
        End With
        .LinkedStyle = ""
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(7)
        .NumberFormat = "%1.%2.%3.%4.%5.%6.%7"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleLegal
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.27)
        .TabPosition = CentimetersToPoints(1.27)
        .ResetOnHigher = 6
        .StartAt = 1
        With .Font
            .Bold = True
            .Italic = False
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdColorDarkRed
            .Size = 10
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = "����"
        End With
        .LinkedStyle = ""
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(8)
        .NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleLegal
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.27)
        .TabPosition = CentimetersToPoints(1.27)
        .ResetOnHigher = 7
        .StartAt = 1
        With .Font
            .Bold = True
            .Italic = False
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdColorDarkRed
            .Size = 10
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = "����"
        End With
        .LinkedStyle = ""
    End With
    With ListGalleries(wdOutlineNumberGallery).ListTemplates(1).ListLevels(9)
        .NumberFormat = "%1.%2.%3.%4.%5.%6.%7.%8.%9"
        .TrailingCharacter = wdTrailingSpace
        .NumberStyle = wdListNumberStyleLegal
        .NumberPosition = CentimetersToPoints(0)
        .Alignment = wdListLevelAlignLeft
        .TextPosition = CentimetersToPoints(1.27)
        .TabPosition = CentimetersToPoints(1.27)
        .ResetOnHigher = 8
        .StartAt = 1
        With .Font
            .Bold = True
            .Italic = False
            .StrikeThrough = wdUndefined
            .Subscript = wdUndefined
            .Superscript = wdUndefined
            .Shadow = wdUndefined
            .Outline = wdUndefined
            .Emboss = wdUndefined
            .Engrave = wdUndefined
            .AllCaps = wdUndefined
            .Hidden = wdUndefined
            .Underline = wdUndefined
            .Color = wdColorDarkRed
            .Size = 10
            .Animation = wdUndefined
            .DoubleStrikeThrough = wdUndefined
            .Name = "����"
        End With
        .LinkedStyle = ""
    End With
    ListGalleries(wdOutlineNumberGallery).ListTemplates(1).Name = ""
    Selection.Range.ListFormat.ApplyListTemplateWithLevel ListTemplate:= _
    ListGalleries(wdOutlineNumberGallery).ListTemplates(1), _
    ContinuePreviousList:=True, ApplyTo:=wdListApplyToWholeList, _
    DefaultListBehavior:=wdWord10ListBehavior
End Sub
