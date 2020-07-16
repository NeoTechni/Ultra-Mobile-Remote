Attribute VB_Name = "EffectsAPI"
Option Explicit
    
Public Const HighliteColor As Long = LCAR_LightBlue

'keyboard
Public Text As String, SelStart As Long, SelLength As Long, Symbols As Boolean, Shift As Boolean, Caps As Boolean, TextID As Long, DefaultText As String, OldText As String
Public Operation As String, timebase As Long, Button As String, DebugMode As Boolean
Private isVisible  As Boolean, isNumber As Boolean, min As Long, max As Long, def As Long

'sensor grid
Public Cx As Double, Cy As Double, pX As Double, pY As Double
Public OldX As Long, OldY As Long, swDown As Boolean, swTolerance As Long, swActive As Boolean, Scrolling As Boolean

Public Type VirtualKey
    Name As String
    vKey As Long
    isMediaKey As Boolean
End Type
Public VkeyCount As Long, VkeyList() As VirtualKey, Mkey As Boolean

Public Function IsValidNumber() As Boolean
    IsValidNumber = Val(Text) <= max And Val(Text) >= min
End Function
Public Function ValidateNumber() As Boolean
    If IsValidNumber Then
        ValidateNumber = True
    Else
        Text = CStr(DefaultText)
    End If
End Function

Public Function AddVkey(vKey As Long, Name As String) As Long
    AddVkey = VkeyCount
    VkeyCount = VkeyCount + 1
    ReDim Preserve VkeyList(VkeyCount)
    With VkeyList(VkeyCount - 1)
        .Name = Name
        .vKey = vKey
        .isMediaKey = Mkey
    End With
End Function

Public Function isShift() As Boolean
    isShift = Shift
End Function

Public Function CursorIsVisible() As Boolean
    Static oldPosition As Long, oldWidth As Long
    CursorIsVisible = (Second(Now) Mod 2 = 0) Or (oldPosition <> SelStart) Or (oldWidth <> SelLength)
    oldPosition = SelStart
    oldWidth = SelLength
End Function

Private Sub Emergency()
    TextID = LCAR_FindLCAR(btntxtbox)
End Sub

Public Function CursorPosition(Optional Position As Long = -1) As Long
    Dim oldsize As Long, Start As Long
    oldsize = Dest.Font.Size
    Emergency
    With LCAR_ButtonList(TextID)
        Dest.Font.Size = .TextSize
        If Position = -1 Then
            Start = Dest.TextWidth(GetSelText)
        Else
            Start = .X + 4
            If Start < 0 Then Start = DestWidth - .X
            If Position > 0 Then Start = Start + Dest.TextWidth(Mid(Text, 1, Position)) ': Debug.Print Mid(Text, 1, Position)
        End If
        CursorPosition = Start
    End With
    Dest.Font.Size = oldsize
End Function



'Alter to use DESTID
Public Sub DrawCursor()
    Dim X As Long, Width As Long, Y As Long
    Emergency
    X = CursorPosition(SelStart)
    If SelLength <> 0 Then
        Width = CursorPosition
        If SelLength < 0 Then Width = -Width
    End If
    With LCAR_ButtonList(TextID)
        If Width <> 0 Then
            If Width > 0 Then
                DrawSquare X, Y + 2, Width, .Height - 1, HighliteColor, HighliteColor
                DrawText X - 1, Y + 2, GetSelText, LCAR_Orange, .TextSize
            Else
                DrawSquare X + Width, Y + 2, Abs(Width) + 1, .Height - 1, HighliteColor, HighliteColor
                DrawText X + Width, Y + 2, GetSelText, LCAR_Orange, .TextSize
            End If
        End If
        
        If CursorIsVisible Then
            DrawSquare X - 1, Y + 2, 3, .Height - 1, LCAR_Orange, LCAR_Orange
            
            'DrawLine X, .Y, 1, .Height, LCAR_Orange
        Else
            LCAR_ButtonList(TextID).IsClean = False
        End If
    End With
End Sub





Public Function KeyboardIsVisible() As Boolean
    KeyboardIsVisible = isVisible
End Function

Private Function btntxtbox() As String
    btntxtbox = IIf(isNumber, "btnnumbox", "btntxtbox")
End Function

Public Sub ShowKeyboard(Default As String, Optional Op As String, Optional NumbersOnly As Boolean, Optional Minimum As Long, Optional Maximum As Long)
    isNumber = NumbersOnly
    Emergency
    Operation = Op
    HideAllGroups IIf(NumbersOnly, 8, 5)
    HideAllLists IIf(NumbersOnly, 6, 2)
    LCAR_ButtonList(preview).Visible = False
    DefaultText = Default
    Text = Default
    SelStart = 0
    SelLength = Len(Default)
    LCAR_SetText btntxtbox, Default, , , , , , , , False
    isVisible = True
    min = Minimum
    max = Maximum
    
    GroupList(9).Visible = True
End Sub

Public Function HideKeyboard() As String
    isVisible = False
    HideAllGroups 3
    HideAllLists 0
    LCAR_ButtonList(preview).Visible = True
    HideKeyboard = Text
    
    GroupList(9).Visible = False
End Function

Public Sub DrawEffects()
    If KeyboardIsVisible Then
        IncrementGrid 0.005 '1
        DrawGridAuto
    End If
End Sub


Public Function ProcessKey(Key As String)
    Static LCARid As Long, SymbolID As Long, Value As Long
    Emergency
    
    If Len(Key) > 0 Then
        If Len(Key) = 1 Then
            SetSelText Key
            If Shift Then
                Shift = False
                If LCARid = 0 Then LCARid = LCAR_FindLCAR("frmbottom", , 6)
                LCAR_Blink LCARid, False
            End If
            If Symbols Then
                Symbols = False
                If SymbolID = 0 Then SymbolID = LCAR_FindLCAR("frmbottom", , 4)
                LCAR_Blink SymbolID, False
            End If
        ElseIf InStr(Key, " ") > 1 Then
            Value = Val(Right(Key, Len(Key) - InStrRev(Key, " ")))
            SelLength = 0
            SelStart = 0
            Select Case LCase(Left(Key, InStr(Key, " ") - 1))
                Case "plus"
                    Value = Val(Text) + Value
                    If max > 0 Then If Value > max Then Value = max
                Case "minus": Value = Val(Text) - Value
                    If Value < min Then Value = min
            End Select
            Text = CStr(Value)
        Else
            Select Case LCase(Key)
                Case "cancel"
                    HideKeyboard
                
                Case "left"
                    If Shift Then
                        SelLength = SelLength - 1
                        If SelLength < 0 Then
                            If SelStart + SelLength < 0 Then SelLength = SelLength + 1
                        End If
                    Else
                        SelLength = 0
                        SelStart = SelStart - 1
                        If SelStart < 0 Then ProcessKey "end"
                    End If
                    
                Case "right"
                    If Shift Then
                        SelLength = SelLength + 1
                        If SelLength > 0 Then
                            If SelStart + SelLength > Len(Text) Then SelLength = SelLength - 1
                        End If
                    Else
                        SelLength = 0
                        SelStart = SelStart + 1
                        If SelStart > Len(Text) Then ProcessKey "home"
                    End If
                    
                Case "home"
                    If Shift Then SelLength = SelStart Else SelLength = 0
                    SelStart = 0
                    
                Case "end"
                    If Shift Then SelLength = -(Len(Text) - SelStart) Else SelLength = 0
                    SelStart = Len(Text)
                    
                Case "delete"
                    If Len(Text) > 0 Then
                        If SelLength = 0 Then
                            SelLength = 1
                        End If
                        SetSelText Empty
                    End If
                    
                Case "backspace"
                    If Len(Text) > 0 Then
                        If SelLength = 0 Then
                            If SelStart > 0 Then
                                SelStart = SelStart - 1
                                SelLength = 1
                            End If
                        End If
                        SetSelText Empty
                    End If
                    
                Case "space"
                    ProcessKey " "
                    
                Case "shift"
                    LCARid = LCAR_FindLCAR("frmbottom", , 6)
                    LCAR_Blink LCARid, Not LCAR_isBlinking(LCARid)
                    Shift = LCAR_isBlinking(LCARid)
                    
            End Select
        End If
    End If
    
    LCAR_SetText LCAR_ButtonList(TextID).Name, Text, , , , , , , , False
End Function

Public Function GetSelText() As String
    If Abs(SelLength) = Len(Text) Then
        GetSelText = Text
    Else
        If SelLength > 0 Then
            GetSelText = Mid(Text, SelStart + 1, SelLength)
        ElseIf SelLength < 0 Then
            GetSelText = Mid(Text, SelStart + SelLength + 1, Abs(SelLength))
        End If
    End If
End Function

Public Sub SetSelText(Key As String)
    Dim LSide As Long, RSide As Long, L As String, R As String
    If SelLength > 0 Then
        LSide = SelStart
        RSide = Len(Text) - SelStart - SelLength
        SelStart = SelStart + Len(Key)
    ElseIf SelLength < 0 Then
        LSide = SelStart + SelLength
        RSide = Len(Text) - SelStart
        SelStart = LSide + Len(Key)
    Else
        LSide = SelStart
        RSide = Len(Text) - SelStart
        SelStart = SelStart + Len(Key)
    End If
    If LSide > 0 Then L = Left(Text, LSide)
    If RSide > 0 Then R = Right(Text, RSide)
    Text = L + Key + R
    SelLength = 0
End Sub

Public Function GetKey(Index As Long) As String
    Dim tempstr As String
    If Symbols Then tempstr = LCARlists(2).ListItems(Index).Side
    If Len(tempstr) = 0 Then
        tempstr = LCARlists(2).ListItems(Index).Text
        If (Shift And Caps) Or ((Not Shift) And (Not Caps)) Then
            tempstr = LCase(tempstr)
        'ElseIf (Shift And Not Caps) Or (Not Shift And Caps) Then
        '    tempstr = UCase(tempstr)
        End If
    End If
    GetKey = tempstr
End Function

Public Sub SetupVKEYS()
    Dim temp As Long
    If VkeyCount = 0 Then
        AddVkey 1, "Left Button"
        AddVkey 2, "Right Button"
        AddVkey 3, "Cancel"
        AddVkey 4, "Middle Button"
        
        AddVkey 8, "Backspace"
        AddVkey 9, "Tab"
        
        AddVkey 12, "Clear"
        AddVkey 13, "Enter"

        AddVkey 16, "Shift"
        AddVkey 17, "Ctrl"
        AddVkey 18, "Alt"
        AddVkey 19, "Pause"
        AddVkey 20, "Caps Lock"
        
        AddVkey 27, "Escape"
        
        AddVkey 32, "Space"
        AddVkey 33, "Page Up"
        AddVkey 34, "Page Down"
        AddVkey 35, "End"
        AddVkey 36, "Home"
        AddVkey 37, "Left"
        AddVkey 38, "Up"
        AddVkey 39, "Right"
        AddVkey 40, "Down"
        
        AddVkey 42, "Print"
        AddVkey 43, "Execute"
        AddVkey 44, "PrintScreen"
        AddVkey 45, "Insert"
        AddVkey 46, "Delete"
        
        For temp = 48 To 57
            AddVkey temp, Chr(temp) '0-9
        Next
        
        For temp = 65 To 90
            AddVkey temp, Chr(temp) 'a-z
        Next
        
        AddVkey 91, "Start"
        AddVkey 92, "Right Start"
        AddVkey 93, "Menu"
        
        For temp = 96 To 105
            AddVkey temp, Chr(temp - 48) 'num pad
        Next
        
        AddVkey 106, "*"
        AddVkey 107, "+"
        
        AddVkey 109, "-"  '
        AddVkey 189, "-"
        
        AddVkey 110, "."
        AddVkey 190, "."
        
        AddVkey 111, "/"
        AddVkey 191, "/"
        
        AddVkey 220, "\"
        'AddVkey 186, ";"
        
        For temp = 112 To 123
            AddVkey temp, "F" & temp - 111
        Next
        
        AddVkey 144, "Num Lock"
        AddVkey 145, "Scroll Lock"
         
        AddVkey 192, "`"
        
        Mkey = True
        AddVkey VK_SLEEP, "Sleep"                       ' &H5F=Computer Sleep key
        AddVkey VK_BROWSER_BACK, "Browser Back"         ' &HA6=Windows 2000/XP: Browser Back key
        AddVkey VK_BROWSER_FORWARD, "Browser Forward"   ' &HA7=Windows 2000/XP: Browser Forward key
        AddVkey VK_BROWSER_REFRESH, "Browser Refresh"   ' &HA8=Windows 2000/XP: Browser Refresh key
        AddVkey VK_BROWSER_STOP, "Browser Stop"         ' &HA9=Windows 2000/XP: Browser Stop key
        AddVkey VK_BROWSER_SEARCH, "Browser Search"     ' &HAA=Windows 2000/XP: Browser Search key
        AddVkey VK_BROWSER_FAVORITES, "Browser Favorites"        ' &HAB=Windows 2000/XP: Browser Favorites key
        AddVkey VK_BROWSER_HOME, "Browser Home"         ' &HAC=Windows 2000/XP: Browser Start and Home key
        AddVkey VK_VOLUME_MUTE, "Volume Mute"           ' &HAD=Windows 2000/XP: Volume Mute key
        AddVkey VK_VOLUME_DOWN, "Volume Down"           ' &HAE =Windows 2000/XP: Volume Down key
        AddVkey VK_VOLUME_UP, "Volume Up"               ' &HAF =Windows 2000/XP: Volume Up key
        AddVkey VK_MEDIA_NEXT_TRACK, "Next Track"       ' &HB0 =Windows 2000/XP: Next Track key
        AddVkey VK_MEDIA_PREV_TRACK, "Prev Track"       ' &HB1 =Windows 2000/XP: Previous Track key
        AddVkey VK_MEDIA_STOP, "Stop"                   ' &HB2 =Windows 2000/XP: Stop Media key
        AddVkey VK_MEDIA_PLAY_PAUSE, "Play/Pause"       ' &HB3 =Windows 2000/XP: Play/Pause Media key
        AddVkey VK_LAUNCH_MAIL, "Launch Mail"           ' &HB4 =Windows 2000/XP: Start Mail key
        AddVkey VK_LAUNCH_MEDIA_SELECT, "Media Select"  ' &HB5 =Windows 2000/XP: Select Media key
        AddVkey VK_LAUNCH_APP1, "Launch App1"           ' &HB6 =Windows 2000/XP: Start Application 1 key
        AddVkey VK_LAUNCH_APP2, "Launch App2"           ' &HB7 =Windows 2000/XP: Start Application 2 key
    End If
End Sub
Public Function vKey2String(vKey As Long, Optional Default As String = "Press Any Key") As String
    Dim temp As Long
    SetupVKEYS
    For temp = 0 To VkeyCount - 1
        If VkeyList(temp).vKey = vKey Then
            vKey2String = VkeyList(temp).Name
            Exit For
        End If
    Next
End Function
Public Function vKeyString2ID(Text As String) As Long
    Dim temp As Long
    SetupVKEYS
    vKeyString2ID = -1
    For temp = 0 To VkeyCount - 1
        If StrComp(Text, VkeyList(temp).Name, vbTextCompare) = 0 Then
            vKeyString2ID = VkeyList(temp).vKey
            Exit For
        End If
    Next
End Function

Public Sub SetupEffects()
    SetupVKEYS
    Cx = 0.5
    Cy = 0.5
    pX = 0.5
    pY = 0.5
End Sub







Public Sub DrawGridAuto()
    Dim Width As Long, Height As Long
    Width = DestWidth - 223
    Height = DestHeight - 338
    DrawSensorGrid 110, 88, Width, Height, Width * Cx, Height * Cy
End Sub

Public Sub IncrementGrid(Optional Speed As Double = 0.05)
    If Cx < pX Then
        Cx = Cx + Speed
        If Cx > pX Then Cx = pX
    ElseIf Cx > pX Then
        Cx = Cx - Speed
        If Cx < pX Then Cx = pX
    End If
    If Cy < pY Then
        Cy = Cy + Speed
        If Cy > pY Then Cy = pY
    ElseIf Cy > pY Then
        Cy = Cy - Speed
        If Cy < pY Then Cy = pY
    End If
    If Cx = pX And Cy = pY Then
        Randomize Timer
        pX = Rnd
        pY = Rnd
    End If
End Sub

Public Sub DrawSensorGrid(X As Long, Y As Long, Width As Long, Height As Long, oX As Long, oY As Long, Optional StartSize As Double = 0.1, Optional Factor As Double = 0.95, Optional Border As Long = 2)
    Dim Cx As Double, cWidth As Double, temp As Long
    Static WasVisible As Boolean
    'Units = 2 ^ Lines
    
    DrawSquare X, Y, Width, Height, vbBlack, vbBlack
    DrawSquare X + Border, Y + Border, Width - Border * 2, Height - Border * 2, vbWhite, IIf(RedAlert, LCAR_Red, LCAR_DarkBlue)
        
    cWidth = StartSize * oX
    Cx = oX + X
    
    Dest.DrawWidth = 3
    DrawLine Cx, Y + 1, 1, Height - 2, vbWhite
    DrawLine X + 1, oY + Y, Width - 2, 1, vbWhite
    Dest.DrawWidth = 1
    
    temp = X + Border
    Do While Cx > temp
        cWidth = cWidth * Factor
        Cx = Cx - cWidth
        If Cx > X Then DrawLine Cx, Y, 1, Height, vbWhite
        If cWidth < 2 Then Cx = 0
    Loop
    
    cWidth = StartSize * (Width - oX)
    Cx = oX + X
    temp = X + Width - Border
    Do While Cx < temp
        cWidth = cWidth * Factor
        Cx = Cx + cWidth
        If Cx < temp Then DrawLine Cx, Y, 1, Height, vbWhite
        If cWidth < 2 Then Cx = X + Width
    Loop
    
    cWidth = StartSize * oY
    Cx = oY + Y
    temp = Y + Border
    Do While Cx > temp
        cWidth = cWidth * Factor
        Cx = Cx - cWidth
        If Cx > temp Then DrawLine X, Cx, Width, 1, vbWhite
        If cWidth < 2 Then Cx = 0
    Loop
    
    cWidth = StartSize * (Height - oY)
    Cx = oY + Y
    temp = Y + Height - Border
    Do While Cx < temp
        cWidth = cWidth * Factor
        Cx = Cx + cWidth
        If Cx < temp Then DrawLine X, Cx, Width, 1, vbWhite
        If cWidth < 2 Then Cx = temp
    Loop
    
    'If Not CursorIsVisible Then
    DrawCursor ' Else LCAR_ButtonList(TextID).IsClean = False
    'If WasVisible <> CursorIsVisible Then
    '    WasVisible = Not WasVisible
    '    DrawCursor
    'End If
End Sub

Public Function IsInSensorSweep(ByVal X As Long, ByVal Y As Long) As Boolean
    Dim Width As Long, Height As Long
    Const Left As Long = 110, Top As Long = 88
    If KeyboardIsVisible Then
         If Rotate Then
            Width = X
            X = Y
            Y = Width
        End If
        Width = DestWidth - 223
        Height = DestHeight - 338
        IsInSensorSweep = isWithin(X, Y, Left, Top, Width, Height)
    End If
End Function
