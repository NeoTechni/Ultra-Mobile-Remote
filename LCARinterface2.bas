Attribute VB_Name = "LCARinterface"
Option Explicit

Public Type location
    x As Long
    y As Long
    Width As Long
    Height As Long
End Type

Public Type LCARButton
    Name As String
    Tag As String
    Group As Long
    
    x As Long
    y As Long
    Width As Long
    Height As Long
    
    OldLoc As location
    Dest As location
    
    LWidth As Long 'BarWidth for elbows
    RWidth As Long 'BarHeight for elbows
    
    IsClean As Boolean
    
    Align As Long '-1 = button, 0 to 3 = elbow
    Text As String 'Button only
    TextAlign As Long
    TextSize As Single
    
    RedAlertHold As Long
    RedAlertCycles As Long
    
    State As Long '-1=blinking 0=mouse up, 1=mouse down
    'Phase As Long
    'Direction As Long 'for blinking
    'EdgeColors(ColorSteps) As OLE_COLOR
    PriColor As OLE_COLOR
    SecColor As OLE_COLOR
    
    Visible As Boolean
    Icon As Long
    Enabled As Boolean
    
    ColorID As Long
    
    'GNDN
    Opacity As Long
    DestID As Long
End Type

Public Type ListItem
    Text As String
    Side As String
    
    Tag As String
    Icon As Long
    Selected As Boolean
    Size As Long
    FILEsize As String
    FileLCAR As String
    IsClean As Boolean
    
    color As OLE_COLOR
    ColorID As Long
    LightColor As OLE_COLOR
    
    WhiteSpace As Long
End Type
Public Type List
    ListCount As Long
    ListItems() As ListItem
    
    ColsPortrait As Long
    ColsLandscape As Long
    
    Start As Long
    x As Long
    y As Long
    Width As Long
    Height As Long
    
    ShowSize As Boolean
    
    Name As String
    IsClean As Boolean
    MultiSelect As Boolean
    
    SelectedItems As Long
    SelectedItem As Long
    TotalSize As Long
    
    isDown As Boolean
    Visible As Boolean
    
    RedX As Long
    RedY As Long
    
    WhiteSpace As Long
    LWidth As Long
    RWidth As Long
    SideWidth As Long
    
    TextSize As Single
    
    'GNDN
    Opacity As Long
    DestID As Long
End Type

Public Type LCARGroup
    Visible As Boolean
    RedAlert As Long
    LCARlist() As Long
    LCARcount As Long
End Type

Public Const SpeedReduction As Single = 0.9

Public OldInc As Long, ThreeDmode As Boolean
Public ClickedAtX As Long, OldClickedAtX As Long, oldsize As Long, SizeMode As Boolean, Inertia As Boolean, Speed As Long
Public GroupList() As LCARGroup, GroupCount As Long, GroupsEnumerated As Boolean, ClickedSide As Boolean
Public LCARlists() As List, LCARListCount As Long, isDown As Boolean, ListId As Long, RedAlert As Boolean
Public LCAR_ButtonList() As LCARButton, LCAR_ButtonCount As Long, IsClean As Boolean, State As Boolean

Public Type Destination
    Dest As Form
    Name As String
    'Rotation state
    isRotated As Boolean
    hPrevFont As Long
    hFont As Long
    issetup As Boolean
    F As LOGFONT
    FontName As String
End Type
Public DestList() As Destination, DestCount As Long, CurrentDest As Long
Public Function LCAR_AddDestination(Destination As Form, Optional Name As String) As Long
    If DestCount = 0 Then Set Dest = Destination
    
    LCAR_AddDestination = DestCount
    DestCount = DestCount + 1
    ReDim Preserve DestList(DestCount)
    With DestList(DestCount - 1)
        Set .Dest = Destination
        If Len(Name) = 0 Then Name = Destination.Name
        .Name = LCase(Name)
    End With
End Function
Public Function LCAR_SwitchToDestination(DestID As Long) As Boolean
    If CurrentDest = DestID Then Exit Function
    With DestList(CurrentDest)
        .isRotated = isRotated
        .hPrevFont = hPrevFont
        .hFont = hFont
        .issetup = issetup
        .F = F
        .FontName = FontName
    End With
    With DestList(DestID)
        isRotated = .isRotated
        hPrevFont = .hPrevFont
        hFont = .hFont
        issetup = .issetup
        F = .F
        FontName = .FontName
    End With
    'Set Dest = DestList(DestID)
    LCAR_SwitchToDestination = True
End Function


Public Sub SetupUImode(Mode As String)
    Select Case UCase(Mode) '
        Case "CLASSIC+AA"
            AntiAliasing = True
            ThreeDmode = False
        Case "NEMESIS"
            AntiAliasing = False
            ThreeDmode = True
        Case Else
            Mode = "Classic"
            AntiAliasing = False
            ThreeDmode = False
    End Select
    SaveSetting "LCAR", "MAIN", "UI", Mode
    IsClean = False
End Sub

Public Sub LCAR_FontIncrement(Increment As Long)
    Dim Delta As Long, temp As Long
    Delta = Increment - OldInc
    For temp = 0 To LCAR_ButtonCount - 1
        With LCAR_ButtonList(temp)
            .TextSize = .TextSize + Delta
            .IsClean = False
        End With
    Next
    For temp = 0 To LCARListCount - 1
        With LCARlists(temp)
            .TextSize = .TextSize + Delta
        End With
    Next
    OldInc = Increment
End Sub

Private Function AddGroup() As Long
    AddGroup = GroupCount
    GroupCount = GroupCount + 1
    ReDim Preserve GroupList(GroupCount)
    GroupList(GroupCount - 1).Visible = True
End Function
Private Function ForceGroupCount(Count As Long) As Long
    Dim temp As Long
    For temp = GroupCount To Count + 1
        AddGroup
    Next
    ForceGroupCount = GroupCount
End Function
Private Function AddLCARtoGroup(LCARid As Long, Group As Long) As Long
    With GroupList(Group)
        AddLCARtoGroup = .LCARcount
        .LCARcount = .LCARcount + 1
        ReDim Preserve .LCARlist(.LCARcount)
        .LCARlist(.LCARcount - 1) = LCARid
    End With
End Function

Public Sub SetRedAlert(Optional Enabled As Boolean = True)
    Dim temp As Long
    RedAlert = Enabled
    If Enabled Then
        If Not GroupsEnumerated Then
            're-enumerate LCAR groups (in case things were deleted)
            For temp = 0 To GroupCount - 1
                With GroupList(temp)
                    .LCARcount = 0
                    ReDim .LCARlist(0)
                End With
            Next
    
            For temp = 0 To LCAR_ButtonCount - 1
                AddLCARtoGroup temp, LCAR_ButtonList(temp).Group
            Next
            GroupsEnumerated = True
        End If
    End If
    IsClean = False
End Sub

Public Function LCAR_ListID(Name As String) As Long
    Dim temp As Long
    LCAR_ListID = -1
    For temp = 0 To LCARListCount - 1
        If StrComp(Name, LCARlists(temp).Name, vbTextCompare) = 0 Then
            LCAR_ListID = temp
            Exit For
        End If
    Next
End Function

Public Function LCAR_AddList(Name As String, ColsPortrait As Long, ColsLandscape As Long, x As Long, y As Long, Width As Long, Height As Long, Optional Visible As Boolean = True, Optional WhiteSpace As Long = 3, Optional LWidth As Long = 20, Optional RWidth As Long, Optional SideWidth As Long = 30, Optional ShowSize As Boolean = True) As Long
    LCAR_AddList = LCARListCount
    LCARListCount = LCARListCount + 1
    ReDim Preserve LCARlists(LCARListCount)
    With LCARlists(LCARListCount - 1)
        .ColsPortrait = ColsPortrait
        .ColsLandscape = ColsLandscape
        
        .x = x
        .y = y
        .Width = Width
        .Height = Height
        
        .ShowSize = ShowSize
        
        .Name = Name
        
        .SelectedItem = -1
        .Visible = Visible
        
        .WhiteSpace = WhiteSpace
        .LWidth = LWidth
        .RWidth = RWidth
        .SideWidth = SideWidth
        
        .TextSize = 10
    End With
End Function

Public Function LCAR_TextWidth(Text As String, Size As Long) As Long
    Dim oldsize As Long
    oldsize = Dest.Font.Size
    Dest.Font.Size = Size
    LCAR_TextWidth = Dest.TextWidth(Text)
    Dest.Font.Size = oldsize
End Function

Public Sub LCAR_DrawLists()
    Const ItemHeight As Long = 21, WhiteSpace As Long = 3 ', SizeWidth As Long = 30
    Dim temp As Long, temp2 As Long, temp3 As Long, temp4 As Long, x As Long, y As Long
    Dim ItemsOnScreen As Long, ItemsPerCol As Long, ItemWidth As Long, Cols As Long, color As Long
    Dim Width As Long, Height As Long, tX As Long, tY As Long, SizeWidth As Long, min As Long
    Dim WhiteSpace2 As Long, RText As String
    
    For temp = 0 To LCARListCount - 1
        With LCARlists(temp)
            If .Visible Then
            SizeWidth = .SideWidth
            
            tX = .x
            tY = .y
            Width = .Width
            Height = .Height
            
        
            If tX < 0 Then tX = DestWidth + tX
            If tY < 0 Then tY = DestHeight + tY
            If Width <= 0 Then Width = DestWidth + Width
            If Height <= 0 Then Height = DestHeight + Height
        
            ItemsOnScreen = Height \ (ItemHeight + WhiteSpace)
            Cols = IIf(Rotate, .ColsPortrait, .ColsLandscape)
            ItemWidth = (Width \ Cols) - WhiteSpace
            ItemsPerCol = .ListCount \ Cols
            
            'If ItemsPerCol > ItemsOnScreen Then ItemsPerCol = ItemsOnScreen
            If .ListCount Mod Cols > 0 Then ItemsPerCol = ItemsPerCol + 1
            min = ItemsOnScreen
            If min > ItemsPerCol Then min = ItemsPerCol
            
            x = tX
            For temp2 = 0 To Cols - 1
                temp4 = .Start + (ItemsPerCol * temp2)
                y = tY
                For temp3 = 1 To min
                    If temp4 < .ListCount And temp4 > -1 Then
                        color = .ListItems(temp4).color
                        WhiteSpace2 = .ListItems(temp4).WhiteSpace
                        If SizeMode Then WhiteSpace2 = 41 ' LCAR_TextWidth("0000", ItemHeight) 'FIX THIS!
            
                        If .ListItems(temp4).Selected And State Then color = .ListItems(temp4).LightColor
                        If RedAlert Then
                            color = LCAR_Red
                            If temp2 = .RedX And (temp3 - 1) = .RedY Then color = LCAR_White
                        End If
                        
                        If Len(.ListItems(temp4).Side) = 0 And Len(.ListItems(temp4).FILEsize) = 0 Then   '.Size = -1 Then
                            DrawLCARButton x, y, ItemWidth, ItemHeight, .ListItems(temp4).Text, color, color, .LWidth, .RWidth, WhiteSpace, , .TextSize      'ammend
                        Else
                            RText = .ListItems(temp4).FILEsize
                            If Len(RText) = 0 Then WhiteSpace2 = WhiteSpace
                            If Len(.ListItems(temp4).FileLCAR) > 4 Then WhiteSpace2 = WhiteSpace2 * 2
                            
                            DrawLCARButton x, y, ItemWidth - SizeWidth - .ListItems(temp4).WhiteSpace, ItemHeight, .ListItems(temp4).Text, color, color, .LWidth, .RWidth, WhiteSpace2, , .TextSize  'ammend
                            'DrawSquare X + ItemWidth - SizeWidth, Y, SizeWidth, ItemHeight, Color, Color
                            
                            If SizeMode Or Len(RText) = 0 Then RText = .ListItems(temp4).Side
                            If SizeMode And Len(.ListItems(temp4).FileLCAR) > 0 Then
                                DrawText x + .LWidth, y - 5, .ListItems(temp4).FileLCAR, color, CSng(ItemHeight + 1)
                            Else
                                If Len(.ListItems(temp4).FILEsize) = 0 Then
                                    If SizeWidth = 0 Then
                                        DrawText x + ItemWidth - (WhiteSpace * 3) - .ListItems(temp4).WhiteSpace - Dest.TextWidth(.ListItems(temp4).Side) / 2, y + 2, .ListItems(temp4).Side, vbBlack, .TextSize
                                    Else
                                        DrawText x + ItemWidth + WhiteSpace - SizeWidth / 2 - .ListItems(temp4).WhiteSpace - Dest.TextWidth(.ListItems(temp4).Side) / 2, y + 2, .ListItems(temp4).Side, vbBlack, .TextSize
                                    End If
                                Else
                                    DrawText x + ItemWidth - SizeWidth - .ListItems(temp4).WhiteSpace - Dest.TextWidth(.ListItems(temp4).Side) - WhiteSpace, y + 2, .ListItems(temp4).Side, vbBlack, .TextSize
                                End If
                            End If
                            If SizeWidth > 0 Then DrawLCARButton x + ItemWidth - SizeWidth, y, SizeWidth, ItemHeight, RText, color, color, 0, 0, 0, 5, .TextSize
                            
                            
                        End If
                        y = y + ItemHeight + WhiteSpace
                        temp4 = temp4 + 1
                    End If
                Next
                x = x + ItemWidth + WhiteSpace
            Next
            .IsClean = True
            
            End If
        End With
    Next
End Sub
Public Function LCAR_isBlinking(LCARid As Long) As Boolean
    LCAR_isBlinking = LCAR_ButtonList(LCARid).State = -1
End Function
Public Function LCAR_Blink(LCARid As Long, Optional Blink As Boolean = True)
    With LCAR_ButtonList(LCARid)
        .State = Val(IIf(Blink, -1, 0))
        .IsClean = False
    End With
End Function

Public Sub LCAR_AddListItems(ListId As Long, ParamArray Items() As Variant)
    Dim temp As Long
    For temp = 0 To UBound(Items)
        LCAR_AddListItem ListId, CStr(Items(temp))
    Next
End Sub
Public Function LCAR_AddListItem(ListId As Long, Text As String, Optional color As Long = -1, Optional LightColor As Long = -1, Optional Size As Long = -1, Optional Tag As String, Optional Icon As Long = -1, Optional Selected As Boolean, Optional Side As String, Optional WhiteSpace As Long = -1, Optional FILEsize As String, Optional LCARtext As String) As Long
    LCAR_AddListItem = LCARlists(ListId).ListCount
    LCARlists(ListId).ListCount = LCARlists(ListId).ListCount + 1
    ReDim Preserve LCARlists(ListId).ListItems(LCARlists(ListId).ListCount)
    With LCARlists(ListId).ListItems(LCARlists(ListId).ListCount - 1)
        .Text = UCase(Text)
        .Side = UCase(Side)
        .Tag = Tag
        .Icon = Icon
        .Selected = Selected
        .Size = Size
        .FileLCAR = LCARtext
        If WhiteSpace = -1 Then .WhiteSpace = LCARlists(ListId).WhiteSpace Else .WhiteSpace = WhiteSpace
        
        If Size = -1 And LCARlists(ListId).ShowSize Then
            If color = -1 Then .color = LCAR_LightBlue Else .color = color
            .FILEsize = FILEsize
        Else
            If color = -1 And LCARlists(ListId).ShowSize Then color = SizeToColor(Size)
            .color = color
            If Len(LCARtext) > 0 Then
                .FILEsize = FILEsize
            ElseIf LCARlists(ListId).ShowSize Then
                .FILEsize = SizeToText(Size, " Q", " K", "M", " G")
                .FileLCAR = Format(SizeToLCAR(Size), "0000")
            End If
            'If InStr(.FileSize, ".") = 0 Then .FileSize = Left(.FileSize, Len(.FileSize) - 1) & " " & Right(.FileSize, 1)
        End If
        If LightColor = -1 Then LightColor = AlterBrightness(.color, Brightness)
        .LightColor = LightColor
        .ColorID = LCAR_ColorIDfromColor(.color)
    End With
End Function

Public Sub LCAR_ClearList(ListId As Long, Optional DownToItem As Long)
    If DownToItem = 0 Then
        ReDim LCARlists(ListId).ListItems(0)
    Else
        ReDim Preserve LCARlists(ListId).ListItems(DownToItem)
    End If
    With LCARlists(ListId)
        .Start = 0
        .ListCount = DownToItem
        .IsClean = False
        .SelectedItem = -1
        .SelectedItems = 0
        .TotalSize = 0
    End With
    IsClean = False
End Sub

Public Sub LCAR_AddFolder(ListId As Long, Path As String, Optional Side As String)
    If Len(Path) > 0 Then LCAR_AddListItem ListId, Right(Path, Len(Path) - InStrRev(Path, "\")), LCAR_LightBlue, , , Path & "\", , , , , Side
End Sub

Public Sub LCAR_EnumFiles(ListId As Long, Optional DriveBox As DriveListBox, Optional Dirbox As DirListBox, Optional FILEbox As FileListBox, Optional Path As String, Optional SortBy As Long = 1, Optional SeparateExtention As Boolean = True, Optional HideExtention As Boolean, Optional SearchQuery As String, Optional Pattern As String = "*.*", Optional Side As String)
    'On Error Resume Next
    Dim temp As Long, File As String, Extention As String, Text As String, AddFile As Boolean
    Dim FileList() As String, FileCount As Long
    
    If Len(Path) = 0 Then
        For temp = 0 To DriveBox.ListCount - 1
            File = DriveBox.List(temp)
            If InStr(File, "[") = 0 Then
                File = File & " [" & DriveType(File) & "]"
            Else
                File = File & " " & FormatPercent(GetPercentFull(File), 2) & " of " & Round(GetTotalSpaceGigaBytes(File), 2) & " Gigaquads used"
            End If
            LCAR_AddListItem ListId, File, LCAR_LightBlue, , , UCase(Left(DriveBox.List(temp), 2)) & "\" ', , , "Drive"
        Next
        
        'LCAR_AddFolder ListId, ShellFolder("Desktop")
        'LCAR_AddFolder ListId, ShellFolder
        'LCAR_AddFolder ListId, ShellFolder("My Music")
        'LCAR_AddFolder ListId, ShellFolder("My Pictures")
        'LCAR_AddFolder ListId, ShellFolder("My Video")
        
        'API_ListBookmarks ListId
        
    Else
        If InStr(Path, "\\") Then Exit Sub
        
        If Not Dirbox Is Nothing Then
        If StrComp(Dirbox.Path, Path, vbTextCompare) = 0 Then
            Dirbox.Refresh
        Else
            If direxists(Path) Then Dirbox.Path = Path Else Exit Sub
        End If
        For temp = 0 To Dirbox.ListCount - 1
            AddFile = True
            Text = Right(Dirbox.List(temp), Len(Dirbox.List(temp)) - InStrRev(Dirbox.List(temp), "\"))
            If Len(SearchQuery) > 0 Then AddFile = SearchText(Text, SearchQuery)
            If InStr(Text, "?") > 0 Then AddFile = False
            If AddFile Then LCAR_AddListItem ListId, Text, -1, -1, -1, Dirbox.List(temp), -1, False ', "Folder"
        Next
        End If
        
        
        If Not FILEbox Is Nothing Then
            FILEbox.Pattern = Pattern
            If StrComp(FILEbox.Path, Path, vbTextCompare) = 0 Then
                FILEbox.Refresh
            Else
                FILEbox.Path = Path
            End If
            FileCount = EnumSortedFiles(FILEbox, FileList, SortBy)
            For temp = 0 To FileCount - 1 'FileBox.ListCount - 1
                AddFile = True
                File = FileList(0, temp) ' Replace(FileBox.Path & "\" & FileBox.List(temp), "\\", "\")
                Text = Right(File, Len(File) - InStrRev(File, "\")) 'FileBox.List(temp)
                If Len(SearchQuery) > 0 Then AddFile = SearchText(Text, SearchQuery)
                If InStr(Text, "?") > 0 Then AddFile = False
                If AddFile Then
                    If InStr(Text, ".") Then
                        If SeparateExtention Then Extention = GetExtention(Text) ' Right(Text, Len(Text) - InStrRev(Text, "."))
                        If SeparateExtention Or HideExtention Then Text = Left(Text, InStrRev(Text, ".") - 1)
                        If Len(Side) > 0 Then Extention = Side
                    End If
                    LCAR_AddListItem ListId, Text, -1, -1, CLng(FileList(2, temp)), File, -1, False, UCase(Extention)    'FileLen(File)
                End If
            Next
        End If
    End If
End Sub

Public Function SearchText(Text As String, SearchQuery As String) As Boolean
    If isapattern(SearchQuery) Then
        SearchText = islike(Text, SearchQuery)
    Else
        SearchText = InStr(1, Text, SearchQuery, vbTextCompare) = 0
    End If
End Function

Public Function isapattern(Text As String) As Boolean
    isapattern = InStr(Text, "?") > 0 Or InStr(Text, "*") > 0
End Function
Public Function islike(Text As String, expression As String) As Boolean 'islike("*.exe", "test.exe")
    Dim tempstr() As String, Count As Long
    expression = LCase(expression)
    Text = LCase(Text)
    If InStr(expression, ";") > 0 Then
        tempstr = Split(expression, ";")
        For Count = 0 To UBound(tempstr)
            If Text Like tempstr(Count) Then
                islike = True
                Exit For
            End If
        Next
    Else
        islike = Text Like expression
    End If
End Function

Public Function SizeToColor(Size As Long) As Long
    Select Case Size 'FileLen(File)
        Case 0 To 1024: SizeToColor = LCAR_Orange '0 to 1 KB
        Case 1025 To 13107: SizeToColor = LCAR_Purple '1 KB to 12.8 KB
        Case 13108 To 1048576: SizeToColor = LCAR_Yellow '12.8 KB to 1 MB
        Case 1048577 To 13421772: SizeToColor = LCAR_DarkBlue '1 MB to 12.8 MB
        Case 13421773 To 1073741824: SizeToColor = LCAR_DarkYellow ' 12.8 MB to 128 MB
        Case Else: SizeToColor = LCAR_DarkPurple '128 MB to infinite
    End Select
End Function

Public Sub LCAR_ANISetLCARLocs(Name As String, x As Long, y As Long, ParamArray Indexes() As Variant)
    Dim temp As Long, LCARid As Long
    For temp = 0 To UBound(Indexes)
        'LCARid = LCAR_FindLCAR(Name, , Indexes(temp))
        If LCARid > -1 Then LCAR_ANIMoveLCAR LCARid, x, y, , , True
    Next
End Sub
Public Sub LCAR_ANISetLCARLoc(LCARid As Long, Optional x As Long = -999, Optional y As Long = -999, Optional Width As Long, Optional Height As Long)
    With LCAR_ButtonList(LCARid)
        If x > -999 Then .x = x
        .OldLoc.x = .x
        
        If y > -999 Then .y = y
        .OldLoc.y = .y
        
        If Width <> 0 Then .Width = Width
        .OldLoc.Width = .Width
        
        If Height <> 0 Then .Height = Height
        .OldLoc.Height = .Height
    End With
End Sub
Public Sub LCAR_ANIMoveLCAR(LCARid As Long, x As Long, y As Long, Optional Width As Long, Optional Height As Long, Optional Relative As Boolean)
    With LCAR_ButtonList(LCARid)
        If Relative Then
            .Dest.x = .x + x
            .Dest.y = .y + y
            .Dest.Width = .Width + Width
            .Dest.Height = .Height + Height
        Else
            .Dest.x = x
            .Dest.y = y
            If Width = 0 Then Dest.Width = .Width Else .Dest.Width = Width
            If Height = 0 Then Dest.Height = .Height Else .Dest.Height = Height
        End If
    End With
End Sub

Public Sub LCAR_ANIIncrementLocs(Optional ReturnToOld As Boolean, Optional Speed As Long = 10, Optional LockAll As Boolean)
    Dim temp As Long
    For temp = 0 To LCAR_ButtonCount - 1
        With LCAR_ButtonList(temp)
            If ReturnToOld Then
                .x = .OldLoc.x
                .y = .OldLoc.y
                .Width = .OldLoc.Width
                .Height = .OldLoc.Height
            ElseIf LockAll Then
                .OldLoc.x = .x
                .OldLoc.y = .y
                .OldLoc.Width = .Width
                .OldLoc.Height = .Height
            Else
                .x = LCAR_Increment(.x, .Dest.x, Speed)
                .y = LCAR_Increment(.y, .Dest.y, Speed)
                .Width = LCAR_Increment(.Width, .Dest.Width, Speed)
                .Height = LCAR_Increment(.Height, .Dest.Height, Speed)
            End If
        End With
    Next
End Sub

Public Function LCAR_Increment(ByVal Current As Long, Destination As Long, Speed As Long) As Long
    LCAR_Increment = Destination
    If Current < Destination Then
        If Current + Speed < Destination Then LCAR_Increment = Current + Speed
    ElseIf Current > Destination Then
        If Current - Speed > Destination Then LCAR_Increment = Current - Speed
    End If
End Function

Public Sub RotateScreen()
    Rotate = Not Rotate
    LCAR_DrawLCARs True
End Sub


Public Function LCAR_AddLCAR(Name As String, x As Long, y As Long, Width As Long, Height As Long, LWidth As Long, RWidth As Long, Optional LightColor As OLE_COLOR = LCAR_LightOrange, Optional DarkColor As OLE_COLOR = LCAR_DarkOrange, Optional Align As Long = -1, Optional Text As String, Optional Tag As String, Optional Group As Long, Optional Visible As Boolean = True, Optional Icon As Long = -1, Optional TextAlign As Long = 4, Optional Enabled As Boolean = True, Optional TextSize As Single) As Long
    Dim temp As Long, temp2 As Double, Alpha As Double
    LCAR_AddLCAR = LCAR_ButtonCount
    LCAR_ButtonCount = LCAR_ButtonCount + 1
    ReDim Preserve LCAR_ButtonList(LCAR_ButtonCount)
    With LCAR_ButtonList(LCAR_ButtonCount - 1)
        .Name = Name
        .Tag = Tag
        .Group = Group
        ForceGroupCount Group
        
        .x = x
        .y = y
        .Width = Width
        .Height = Height
        
        .OldLoc.x = x
        .OldLoc.y = y
        .OldLoc.Width = Width
        .OldLoc.Height = Height
        
        .LWidth = LWidth
        .RWidth = RWidth
        
        .Align = Align
        .Text = UCase(Text)
        .TextAlign = TextAlign
        .Visible = Visible
        .Icon = Icon
        If TextSize = 0 Then
            .TextSize = Dest.Font.Size
        Else
            .TextSize = TextSize
        End If
        
        .PriColor = DarkColor
        .ColorID = LCAR_ColorIDfromColor(DarkColor)
        .SecColor = LightColor
        If LightColor = -1 Then .SecColor = AlterBrightness(DarkColor, Brightness)
        
        .Enabled = Enabled
        'temp2 = 256 / ColorSteps
        'For temp = 0 To ColorSteps - 1
        '    '.EdgeColors(temp) = AlphaBlend(MidColor, DarkColor, Alpha)
        '    .FillColors(temp) = AlphaBlend(LightColor, DarkColor, Alpha)
        '    Alpha = Alpha + temp2
        'Next
    End With
End Function

Public Sub LCAR_MoveLCARs(Name As String, Start As Long, Optional Count As Long = 1, Optional ByVal x As Long = -1, Optional ByVal y As Long = -1, Optional Group As Long = -1, Optional AlongXaxis As Boolean = True, Optional WhiteSpace As Long = 2, Optional Width As Long, Optional Height As Long)
    Dim temp As Long, LCAR As Long
    For temp = Start To Start + Count - 1
        LCAR = LCAR_FindLCAR(Name, Group, temp)
        If LCAR > -1 Then
            With LCAR_ButtonList(LCAR)
                If x = -1 Then x = .x Else .x = x
                If y = -1 Then y = .y Else .y = y
                If Height <> 0 Then .Height = Height
                If Width <> 0 Then .Width = Width
                If AlongXaxis Then
                    x = x + WhiteSpace + .Width
                Else
                    y = y + WhiteSpace + .Height
                End If
                .OldLoc.x = .x
                .OldLoc.y = .y
                .OldLoc.Width = .Width
                .OldLoc.Height = .Height
            End With
        Else
            Exit For
        End If
    Next
End Sub

Public Sub LCAR_HideLCAR(Name As String, Optional Visible As Boolean)
    Dim temp As Long
    For temp = 0 To LCAR_ButtonCount - 1
        If StrComp(Name, LCAR_ButtonList(temp).Name, vbTextCompare) = 0 Then
            LCAR_ButtonList(temp).Visible = Visible
            LCAR_ButtonList(temp).IsClean = False
            IsClean = False
        End If
    Next
End Sub

Public Function LCAR_FindLCAR(Name As String, Optional Group As Long = -1, Optional Index As Long = 0) As Long 'If Index=-1 then it will count the occurances of that button id
    Dim temp As Long, temp2 As Long
    LCAR_FindLCAR = -1
    For temp = 0 To LCAR_ButtonCount - 1
        If StrComp(Name, LCAR_ButtonList(temp).Name, vbTextCompare) = 0 Then
            If Group = -1 Or Group = LCAR_ButtonList(temp).Group Then
                If Index = 0 Then
                    LCAR_FindLCAR = temp
                    Exit For
                Else
                    If temp2 = Index Then
                        LCAR_FindLCAR = temp
                        Exit For
                    End If
                    temp2 = temp2 + 1
                End If
            End If
        End If
    Next
    If Index = -1 Then LCAR_FindLCAR = temp2
End Function

Public Function LCAR_FindIndex(ButtonID As Long) As Long
    Dim temp As Long, Name As String, Index As Long
    Name = LCAR_ButtonList(ButtonID).Name
    For temp = ButtonID - 1 To 0 Step -1
        If StrComp(Name, LCAR_ButtonList(temp).Name, vbTextCompare) = 0 Then Index = Index + 1
    Next
    LCAR_FindIndex = Index
End Function

Public Function isWithin(x As Long, y As Long, Left As Long, Top As Long, Width As Long, Height As Long) As Boolean
    If x >= Left Then
        If x < Left + Width Then
            If y >= Top Then isWithin = y < Top + Height
        End If
    End If
End Function

Public Sub LCAR_Select(ListId As Long, Operation As Long)
    Dim temp As Long
    With LCARlists(ListId)
        .TotalSize = 0
        For temp = 0 To .ListCount - 1
            Select Case Operation
                Case -1 'invert selection
                    .ListItems(temp).Selected = Not .ListItems(temp).Selected
                Case 0 'select none
                    If .ListItems(temp).Selected Then
                        .ListItems(temp).Selected = False
                        .ListItems(temp).IsClean = False
                    End If
                Case 1 'select all
                    If Not .ListItems(temp).Selected Then
                        .ListItems(temp).Selected = True
                        .ListItems(temp).IsClean = False
                    End If
            End Select
            If .ListItems(temp).Selected And .ListItems(temp).Size > -1 Then .TotalSize = .TotalSize + .ListItems(temp).Size
        Next
        .IsClean = False
        Select Case Operation
            Case -1: .SelectedItems = .ListCount - .SelectedItems
            Case 0: .SelectedItems = 0
            Case 1: .SelectedItems = .ListCount
        End Select
    End With
End Sub

Public Sub LCAR_SelectItem(ListId As Long, ItemID As Long)
    Dim temp As Long
    If ItemID = -1 Then Exit Sub
    With LCARlists(ListId)
        If Not .MultiSelect Then LCAR_Select ListId, 0
        .ListItems(ItemID).Selected = Not .ListItems(ItemID).Selected
        .ListItems(ItemID).IsClean = False
        If .ListItems(ItemID).Selected Then
            .SelectedItems = .SelectedItems + 1
            .SelectedItem = ItemID
            If .ListItems(ItemID).Size > -1 Then .TotalSize = .TotalSize + .ListItems(ItemID).Size
        Else
            If .ListItems(ItemID).Size > -1 Then .TotalSize = .TotalSize - .ListItems(ItemID).Size
            .SelectedItem = -1
            .SelectedItems = .SelectedItems - 1
            If .SelectedItems = 1 Then
                For temp = 0 To .ListCount - 1
                    If .ListItems(temp).Selected Then
                        .SelectedItem = temp
                        Exit For
                    End If
                Next
            End If
        End If
    End With
End Sub

Public Function LCAR_SelectedItem(ListId As Long) As String
    With LCARlists(ListId)
        If .SelectedItem > -1 Then
            LCAR_SelectedItem = .ListItems(.SelectedItem).Text
        End If
    End With
End Function

Public Function LCAR_ListRows(ListId As Long) As Long
    Const ItemHeight As Long = 21, WhiteSpace As Long = 3
    Dim Height As Long
    With LCARlists(ListId)
        Height = .Height
        If Height <= 0 Then Height = DestHeight + Height
        LCAR_ListRows = Height \ (ItemHeight + WhiteSpace)
    End With
End Function
Public Function LCAR_ListCols(ListId As Long) As Long
    With LCARlists(ListId)
        LCAR_ListCols = IIf(Rotate, .ColsPortrait, .ColsLandscape)
    End With
End Function
Public Function LCAR_ListHeight(ListId As Long) As Long
    Const ItemHeight As Long = 21, WhiteSpace As Long = 3
    Dim ItemsOnScreen As Long, ItemsPerCol As Long, ItemWidth As Long, Cols As Long, Height As Long
    With LCARlists(ListId)
        Height = .Height
        If Height <= 0 Then Height = DestHeight + Height
        ItemsOnScreen = Height \ (ItemHeight + WhiteSpace)
        Cols = IIf(Rotate, .ColsPortrait, .ColsLandscape)
        ItemsPerCol = .ListCount \ Cols
        
        If ItemsPerCol < ItemsOnScreen Then
            LCAR_ListHeight = ItemsOnScreen * (ItemHeight + WhiteSpace)
        Else
            LCAR_ListHeight = Height
        End If
    End With
End Function

Public Function LCAR_ClickedCol(ListId As Long, ByVal x As Long, ByVal y As Long, Optional AllowOB As Boolean = True) As Long
    Const ItemHeight As Long = 21, WhiteSpace As Long = 3, SizeWidth As Long = 30
    Dim temp As Long, tX As Long, tY As Long, Width As Long, Height As Long
    Dim ItemsOnScreen As Long, ItemsPerCol As Long, ItemWidth As Long, Cols As Long, color As Long
    LCAR_ClickedCol = -1
    If Rotate Then
        temp = x
        x = Dest.ScaleHeight - y
        y = temp
    End If
    With LCARlists(ListId)
        tX = .x
        Width = .Width
        If tX < 0 Then tX = DestWidth + tX
        If Width <= 0 Then Width = DestWidth + Width
        Cols = IIf(Rotate, .ColsPortrait, .ColsLandscape)
        ItemWidth = (Width \ Cols)
        x = x - tX
        temp = (x \ ItemWidth)
        If AllowOB Then
            LCAR_ClickedCol = temp
        Else
            If temp > -1 And temp < Cols Then LCAR_ClickedCol = temp
        End If
    End With
End Function

Public Function LCAR_ClickedRow(ListId As Long, ByVal x As Long, ByVal y As Long, Optional AllowOB As Boolean = True) As Long
    Const ItemHeight As Long = 21, WhiteSpace As Long = 3 ', SizeWidth As Long = 30
    Dim tY As Long, temp As Long, Cols As Long, ItemsPerCol As Long, Height As Long, ItemsOnScreen As Long
    LCAR_ClickedRow = -1
    If Rotate Then
        temp = x
        x = Dest.ScaleHeight - y
        y = temp
    End If
    With LCARlists(ListId)
        tY = .y
        Height = .Height
        If tY < 0 Then tY = DestHeight + tY
        If Height <= 0 Then Height = DestHeight + Height
        y = y - tY
        temp = y \ (ItemHeight + WhiteSpace)
        Cols = IIf(Rotate, .ColsPortrait, .ColsLandscape)
        ItemsOnScreen = Height \ (ItemHeight + WhiteSpace)
        If AllowOB Then
            LCAR_ClickedRow = temp
        Else
            If temp > -1 And temp < ItemsOnScreen Then LCAR_ClickedRow = temp
        End If
    End With
End Function


Public Sub LCAR_ScrollList(ListId As Long, Rows As Long)
    Const ItemHeight As Long = 21, WhiteSpace As Long = 3
    Dim Cols As Long, ItemsPerCol As Long, Height As Long, ItemsOnScreen As Long, OldStart As Boolean
    With LCARlists(ListId)
        OldStart = .Start
        Height = .Height
        If Height <= 0 Then Height = DestHeight + Height
        ItemsOnScreen = Height \ (ItemHeight + WhiteSpace)
        Cols = IIf(Rotate, .ColsPortrait, .ColsLandscape)
        ItemsPerCol = .ListCount \ Cols
        .Start = .Start + Rows
        If .Start < 0 Then
            .Start = 0
        ElseIf .Start >= ItemsPerCol - ItemsOnScreen Then
            .Start = ItemsPerCol - ItemsOnScreen
        End If
        .IsClean = OldStart = .Start
        If .IsClean Then Exit Sub
    End With
    LCAR_DrawLCARs True
End Sub

Public Function LCARS_ListItemsOnScreen(ListId As Long)
    Const ItemHeight As Long = 21, WhiteSpace As Long = 3
    Dim ItemsOnScreen As Long, ItemsPerCol As Long, Cols As Long, Height As Long
    
    With LCARlists(ListId)
        Height = .Height
        If Height <= 0 Then Height = DestHeight + Height
        ItemsOnScreen = Height \ (ItemHeight + WhiteSpace)
        Cols = IIf(Rotate, .ColsPortrait, .ColsLandscape)
        ItemsPerCol = .ListCount \ Cols
        If ItemsOnScreen < ItemsPerCol Then LCARS_ListItemsOnScreen = ItemsOnScreen Else LCARS_ListItemsOnScreen = ItemsPerCol
    End With
End Function

Public Function LCAR_ListItem(ListId As Long, Text As String, Optional Side As Boolean) As Long
    Dim temp As Long, Found As Boolean
    LCAR_ListItem = -1
    With LCARlists(ListId)
        For temp = 0 To .ListCount - 1
            If Side Then
                Found = StrComp(Text, .ListItems(temp).Tag, vbTextCompare) = 0
            Else
                Found = StrComp(Text, .ListItems(temp).Text, vbTextCompare) = 0
            End If
            If Found Then
                LCAR_ListItem = temp
                Exit For
            End If
        Next
    End With
End Function

'new!
Public Function LCAR_FindListItemByName(ListId As Long, Text As String, Optional ByTag As Boolean) As Long
    Dim temp As Long
    LCAR_FindListItemByName = -1
    For temp = 0 To LCARlists(ListId).ListCount - 1
        With LCARlists(ListId).ListItems(temp)
            If StrComp(Text, IIf(ByTag, .Tag, .Text), vbTextCompare) = 0 Then
                LCAR_FindListItemByName = temp
                Exit For
            End If
        End With
    Next
End Function
Public Function LCAR_FindListItem(ByVal x As Long, ByVal y As Long) As Long
    Const ItemHeight As Long = 21, WhiteSpace As Long = 3, SizeWidth As Long = 30
    Dim temp As Long, tX As Long, tY As Long, Width As Long, Height As Long, ListId As Long, OldX As Long
    Dim ItemsOnScreen As Long, ItemsPerCol As Long, ItemWidth As Long, Cols As Long, color As Long
    
    LCAR_FindListItem = -1
    ListId = LCAR_FindList(x, y)
    If ListId = -1 Then Exit Function

    If Rotate Then
        temp = x
        x = Dest.ScaleHeight - y
        y = temp
    End If
    With LCARlists(ListId)
        tX = .x
        tY = .y
        Width = .Width
        Height = .Height
        If tX < 0 Then tX = DestWidth + tX
        If tY < 0 Then tY = DestHeight + tY
        If Width <= 0 Then Width = DestWidth + Width
        If Height <= 0 Then Height = DestHeight + Height
        
        ItemsOnScreen = Height \ (ItemHeight + WhiteSpace)
        Cols = IIf(Rotate, .ColsPortrait, .ColsLandscape)
        ItemWidth = (Width \ Cols)
        ItemsPerCol = .ListCount \ Cols
        If .ListCount Mod Cols > 0 Then ItemsPerCol = ItemsPerCol + 1
        
        x = x - tX
        y = y - tY
        OldX = x Mod ItemWidth
        
        x = (x \ ItemWidth)
        y = y \ (ItemHeight + WhiteSpace)
        
        If y < ItemsPerCol Then
            y = y + .Start + (ItemsPerCol * x)
            temp = y
            ClickedSide = False
            If Len(.ListItems(temp).Side) > 0 Or Len(.ListItems(temp).FILEsize) > 0 Then ClickedSide = OldX > (ItemWidth - .SideWidth - WhiteSpace * 2)
            If temp < .ListCount Then LCAR_FindListItem = temp
        End If
    End With
End Function
Public Function LCAR_FindList(ByVal x As Long, ByVal y As Long) As Long
    Dim temp As Long, tX As Long, tY As Long, Width As Long, Height As Long
    
    LCAR_FindList = -1
    If Rotate Then
        temp = x
        x = Dest.ScaleHeight - y
        y = temp
    End If
    For temp = LCARListCount - 1 To 0 Step -1
        With LCARlists(temp)
            If .Visible Then
                tX = .x
                tY = .y
                Width = .Width
                Height = .Height
                If tX < 0 Then tX = DestWidth + tX
                If tY < 0 Then tY = DestHeight + tY
                If Width <= 0 Then Width = DestWidth + Width
                If Height <= 0 Then Height = DestHeight + Height
                If isWithin(x, y, tX, tY, Width, Height) Then
                    LCAR_FindList = temp
                    Exit For
                End If
            End If
        End With
    Next
End Function


Public Function LCAR_FindClicked(ByVal x As Long, ByVal y As Long, Optional IncludeElbows As Boolean) As Long
    Dim temp As Long, Found As Boolean, tX As Long, tY As Long, Width As Long, Height As Long
    LCAR_FindClicked = -1
    If Rotate Then
        'RotateXY X, Y
        temp = x
        x = Dest.ScaleHeight - y
        y = temp
    End If
    For temp = LCAR_ButtonCount - 1 To 0 Step -1 ' reverse order so those drawn on top get clicked first!
        With LCAR_ButtonList(temp)
            If .Visible And GroupList(.Group).Visible Then
                tX = .x
                tY = .y
                Width = .Width
                Height = .Height
                If tX < 0 Then tX = DestWidth + tX
                If tY < 0 Then tY = DestHeight + tY
                If Width <= 0 Then Width = DestWidth + Width
                If Height <= 0 Then Height = DestHeight + Height
        
                If .Align = -1 Then
                    Found = isWithin(x, y, tX, tY, Width, Height)
                    ClickedAtX = x - tX
                ElseIf IncludeElbows Then
                    Found = isWithin(x, y, tX, tY, Width, Height)
                    If Found Then
                        Select Case .Align
                            Case 0: Found = Not isWithin(x, y, tX + .LWidth, tY + .RWidth, Width - .LWidth, Height - .RWidth) '|-  top left
                            Case 1: Found = Not isWithin(x, y, tX, tY + .RWidth, Width - .LWidth, Height - .RWidth) '-| top right
                            Case 2: Found = Not isWithin(x, y, tX + .LWidth, tY, Width - .LWidth, Height - .RWidth) '|_ bottom left
                            Case 3: Found = Not isWithin(x, y, tX, tY, Width - .LWidth, Height - .RWidth)  '_| bottom right
                        End Select
                    End If
                End If
            End If
        End With
        If Found Then
            LCAR_FindClicked = temp
            Exit For
        End If
    Next
End Function

Public Sub LCAR_BlinkLCARs()
    Dim temp As Long, temp2 As Long, looped As Boolean
    If RedAlert Then
        For temp = 0 To GroupCount - 1
            With GroupList(temp)
                If .Visible Then
                    .RedAlert = .RedAlert + 1
                    If .RedAlert >= .LCARcount Then .RedAlert = 0
                    If Not LCAR_ButtonList(.LCARlist(.RedAlert)).Visible Then
                        For temp2 = .RedAlert To .LCARcount - 1
                            .RedAlert = .RedAlert + 1
                            If LCAR_ButtonList(.LCARlist(.RedAlert)).Visible Then Exit For
                        Next
                    End If
                End If
            End With
        Next
        For temp = 0 To LCARListCount - 1
            With LCARlists(temp)
                If .Visible Then
                    .RedY = .RedY + 1
                    If .RedY > LCAR_ListRows(temp) Or .RedY > LCARS_ListItemsOnScreen(temp) Then
                        .RedY = 0
                        .RedX = .RedX + 1
                        If .RedX = LCAR_ListCols(temp) Then .RedX = 0
                    End If
                    .IsClean = False
                End If
            End With
        Next
        IsClean = False
    Else
        State = Not State
        For temp = 0 To LCAR_ButtonCount - 1
            If LCAR_ButtonList(temp).State = -1 Then LCAR_ButtonList(temp).IsClean = False
        Next
    End If
    LCAR_DrawLCARs
End Sub

Public Sub LCAR_DrawLCARs(Optional ClearScreen As Boolean)
    Dim temp As Long, x As Long, y As Long, Width As Long, Height As Long, color As OLE_COLOR
    Dim EdgePen As Long, FillBrush As Long, OldEdge As Long, OldFill As Long, TextColor As OLE_COLOR
    TextColor = vbBlack
    If ClearScreen Or Not IsClean Then
        IsClean = False
        'find a better way to clear!
        
        'Method 1 fails
        'Dest.BackColor = vbBlack
                
        'Method 2 draws a random color
        'Dest.FillStyle = vbSolid
        'Dest.Line (0, 0)-(200, 200), vbBlack, B '   (Dest.ScaleWidth, Dest.ScaleHeight), vbBlack, B
        'Dest.FillStyle = 1
        
        'Method 3 fails
        'SwitchToUnRotated
        'Dest.Cls
        
        'Method 4 fails
        'EdgePen = CreatePen(PS_SOLID, 15, vbBlack)
        'DeleteObject SelectObject(Dest.hdc, EdgePen)
        'FillBrush = CreateSolidBrush(vbBlack)
        'DeleteObject SelectObject(Dest.hdc, FillBrush)
        'RectangleX Dest.hdc, 0, 0, Dest.ScaleWidth, Dest.ScaleHeight
        
        EdgePen = CreatePen(PS_SOLID, 15, vbBlack)
        OldEdge = SelectObject(Dest.hDC, EdgePen)
        FillBrush = CreateSolidBrush(vbBlack)
        OldFill = SelectObject(Dest.hDC, FillBrush)
        RectangleX Dest.hDC, 0, 0, Dest.ScaleWidth, Dest.ScaleHeight
        
        SelectObject Dest.hDC, OldEdge
        SelectObject Dest.hDC, OldFill
        DeleteObject EdgePen
        DeleteObject FillBrush
    End If
    If Dest.WindowState = vbMinimized Then Exit Sub
    For temp = 0 To LCAR_ButtonCount - 1
        With LCAR_ButtonList(temp)
            'If temp = 42 Then MsgBox "HI"
            
            If .Visible And GroupList(.Group).Visible And (Not .IsClean Or Not IsClean) Then
                x = .x
                y = .y
                Width = .Width
                Height = .Height

                If x < 0 Then x = DestWidth + x
                If y < 0 Then y = DestHeight + y
                If Width <= 0 Then Width = DestWidth + Width
                If Height <= 0 Then Height = DestHeight + Height
                
                Select Case .State
                    Case -1: If State Then color = .SecColor Else color = .PriColor 'Blinking
                    Case 0: color = .PriColor 'off
                    Case 1: color = .SecColor 'on/mousedown
                End Select
                If RedAlert Then
                    If color = vbBlack Then
                        TextColor = LCAR_Red
                        If GroupList(.Group).LCARlist(GroupList(.Group).RedAlert) = temp Then TextColor = LCAR_White
                    Else
                        color = LCAR_Red
                        If GroupList(.Group).LCARlist(GroupList(.Group).RedAlert) = temp Then color = LCAR_White
                    End If
                End If
                
                If .Align = -1 Then
                    DrawLCARButton x, y, Width, Height, .Text, color, color, .LWidth, .RWidth, , .TextAlign, .TextSize, TextColor, .ColorID
                    TextColor = vbBlack
                Else
                    If .TextSize > 0 Then
                        SwitchToUnRotated
                        Dest.Font.Size = .TextSize
                    End If
                    DrawLCARelbow x, y, Width, Height, .LWidth, .RWidth, , .Align, color, color, .Text, .TextAlign, .ColorID
                End If
            End If
            .IsClean = True
            
            If Dest.Font.Size <> oldsize Then
                SwitchToUnRotated
                Dest.Font.Size = oldsize
            End If
                
        End With
    Next
    LCAR_DrawLists
    
    If ClearScreen Or Not IsClean Then DrawEffects
    
    Dest.Refresh
    IsClean = True
End Sub

Public Sub LCAR_DeleteLCAR(Index As Long)
    Dim temp As Long
    For temp = Index + 1 To LCAR_ButtonCount - 1
        LCAR_ButtonList(temp - 1) = LCAR_ButtonList(temp)
    Next
    LCAR_ButtonCount = LCAR_ButtonCount - 1
    GroupsEnumerated = False
    If LCAR_ButtonCount > 0 Then ReDim Preserve LCAR_ButtonList(LCAR_ButtonCount)
End Sub

Public Sub LCAR_DeleteName(Name As String)
    Dim temp As Long, tempstr As String
    For temp = LCAR_ButtonCount - 1 To 0 Step -1
        tempstr = LCAR_ButtonList(temp).Name
        If StrComp(Name, tempstr, vbTextCompare) = 0 Then LCAR_DeleteLCAR temp
    Next
End Sub
Public Sub LCAR_DeleteGroup(Index As Long)
    Dim temp As Long, temp2 As Long
    If Index > -1 Then
        For temp = LCAR_ButtonCount - 1 To 0 Step -1
            temp2 = LCAR_ButtonList(temp).Group
            If temp2 = Index Then LCAR_DeleteLCAR temp
        Next
    End If
End Sub

Public Sub LCAR_DeleteListItem(ListId As Long, ItemID As Long)
    Dim temp As Long
    For temp = ItemID + 1 To LCARlists(ListId).ListCount - 1
        LCARlists(ListId).ListItems(temp - 1) = LCARlists(ListId).ListItems(temp)
    Next
    LCARlists(ListId).ListCount = LCARlists(ListId).ListCount - 1
    If LCARlists(ListId).ListCount = 0 Then
        ReDim LCARlists(ListId).ListItems(0)
    Else
        ReDim Preserve LCARlists(ListId).ListItems(LCARlists(ListId).ListCount)
    End If
    LCARlists(ListId).IsClean = False
End Sub

Public Sub LCAR_AddMenu(Name As String, Group As Long, x As Long, y As Long, Items As Long, Width As Long, Height As Long, Optional Xaxis As Boolean = True, Optional LWidth As Long, Optional RWidth As Long, Optional WhiteSpace As Long = 2, Optional color As Long = 27607)
    Dim temp As Long
    For temp = 1 To Items
        LCAR_AddLCAR Name, x, y, Width, Height, LWidth, RWidth, AlterBrightness(color, Brightness), color, , , , Group
        If Xaxis Then
            x = x + Width + WhiteSpace
        Else
            y = y + Height + WhiteSpace
        End If
    Next
End Sub

Public Function LCAR_MinWidth(ID As Long) As Long
    Dim temp As Long
    Const SPACE As Long = 4
    temp = LCAR_ButtonList(ID).LWidth
    If temp > 0 Then temp = temp + SPACE
    temp = temp + LCAR_ButtonList(ID).RWidth
    If LCAR_ButtonList(ID).RWidth > 0 Then temp = temp + SPACE
    LCAR_MinWidth = Dest.TextWidth(LCAR_ButtonList(ID).Text) + temp + SPACE + 2
End Function

Public Function LCAR_SetText(Name As String, Text As String, Optional Index As Long, Optional Group As Long = -1, Optional Crop As Boolean, Optional CropAll As Boolean, Optional WhiteSpace As Long = 2, Optional MakeVisible As Boolean = True, Optional Tag As String, Optional UpperCase As Boolean = True) As Boolean
    Dim temp As Long, Count As Long, temp2 As Long, temp3 As Long
    temp = LCAR_FindLCAR(Name, Group, Index)
    If temp > -1 Then
        If UpperCase Then LCAR_ButtonList(temp).Text = UCase(Text) Else LCAR_ButtonList(temp).Text = Text
        LCAR_ButtonList(temp).Tag = Tag
        LCAR_ButtonList(temp).IsClean = False
        LCAR_ButtonList(temp).Enabled = Len(Text) > 0
        'If Len(Side) > 0 Then LCAR_ButtonList(temp).Side = Side
        LCAR_SetText = True
        If Crop Then
            LCAR_ButtonList(temp).Width = LCAR_MinWidth(temp)
            If MakeVisible Then LCAR_ButtonList(temp).Visible = True
            If CropAll Then
                Count = LCAR_FindLCAR(Name, Group, -1)
                For temp2 = Index + 1 To Count - 1
                    temp3 = LCAR_FindLCAR(Name, Group, temp2)
                    If temp3 = -1 Then
                        Exit For
                    Else
                        LCAR_ButtonList(temp3).x = LCAR_ButtonList(temp).Width + LCAR_ButtonList(temp).x + WhiteSpace
                        temp = temp3
                    End If
                Next
            End If
        End If
    End If
End Function

Public Function LCAR_CountLCARs(Name As String, Optional Group As Long) As Long
    LCAR_CountLCARs = LCAR_FindLCAR(Name, Group, -1)
End Function

Public Function LCAR_SetTexts(Name As String, Group As Long, Crop As Boolean, Hide As Boolean, ParamArray Texts() As Variant) As Long
    Dim temp As Long, Count As Long, temp2 As Long
    Count = LCAR_FindLCAR(Name, Group, -1)
    For temp = 0 To UBound(Texts)
        LCAR_SetText Name, CStr(Texts(temp)), temp, Group, Crop, Crop
    Next
    LCAR_SetTexts = UBound(Texts)
    If Hide Then
        For temp = UBound(Texts) + 1 To Count - 1
            temp2 = LCAR_FindLCAR(Name, Group, temp)
            If temp2 > -1 Then
                LCAR_ButtonList(temp2).Visible = False
            End If
        Next
    End If
End Function
Public Function LCAR_SetTextsArray(Name As String, Group As Long, Crop As Boolean, Hide As Boolean, Texts) As Long
    Dim temp As Long, Count As Long, temp2 As Long, Tag As String, count2 As Long
    Count = LCAR_FindLCAR(Name, Group, -1)
    count2 = UBound(Texts) + 1
    For temp = 0 To UBound(Texts)
        If Len(Texts(temp)) = 0 Then
            count2 = count2 - 1
        Else
            Tag = Tag & CStr(Texts(temp)) & "\"
            LCAR_SetText Name, CStr(Texts(temp)), temp, Group, Crop, Crop, , , Tag
        End If
    Next
    LCAR_SetTextsArray = count2 - 1
    If Hide Then
        For temp = count2 To Count - 1
            temp2 = LCAR_FindLCAR(Name, Group, temp)
            If temp2 > -1 Then
                LCAR_ButtonList(temp2).Visible = False
            End If
        Next
    End If
End Function
Public Function LCAR_NextX(Name As String, Optional Group As Long = -1, Optional Index As Long = 0, Optional WhiteSpace As Long = 2) As Long
    Dim temp As Long
    temp = LCAR_FindLCAR(Name, Group, Index)
    If temp > -1 Then LCAR_NextX = LCAR_ButtonList(temp).Width + LCAR_ButtonList(temp).x + WhiteSpace
End Function
Public Function LCAR_NextY(Name As String, Optional Group As Long = -1, Optional Index As Long = 0, Optional WhiteSpace As Long = 2) As Long
    Dim temp As Long
    temp = LCAR_FindLCAR(Name, Group, Index)
    If temp > -1 Then LCAR_NextY = LCAR_ButtonList(temp).Height + LCAR_ButtonList(temp).y + WhiteSpace
End Function

