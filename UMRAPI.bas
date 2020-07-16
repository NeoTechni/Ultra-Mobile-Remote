Attribute VB_Name = "UMRAPI"
    Option Explicit
    
    Public Enum UMR_MouseAction
        UMR_ScrollUp
        UMR_ScrollLeft
        UMR_ScrollRight
        UMR_ScrollDown
        
        UMR_MoveUp
        UMR_MoveLeft
        UMR_MoveRight
        UMR_MoveDown
        
        UMR_LeftButton
        UMR_MiddleButton
        UMR_RightButton
    End Enum
    Public Enum UMR_ActionType
        UMR_DoNothing
        UMR_Keyboard
        UMR_MediaKey
        UMR_Mouse
    End Enum
    Public Type UMR_Action
        ActionType As UMR_ActionType
        Button As Long
        CTRL As Boolean
        ALT As Boolean
    End Type
    Public UMR_ButtonCache(15) As UMR_Action, OldState As String

    Public UMHini As Hini, CurrentProfile As String, CurrentDir As String, IsInFocus As Boolean
    Public Const HTML_Tasklist As String = "%T%<TR><TD><A HREF='%COMMAND%'>%WINDOWNAME%</A></TD></TR>%N%"
    Public Const HTML_Filelist As String = "%T%<TR><TD><A HREF='%COMMAND%'>%FileTitle%</A></TD><TD>%EXTENTION%</TD><TD>%Filesize%</TD></TD>%N%"
    Public Const HTML_Butnlist As String = "%T%<TR><TD><A HREF='%COMMAND%'>%Name%</A></TD></TR>%N%" 'TD  bgcolor=%COLOR%

Public Function UMR_CacheButton(Button As String, Optional Profile As String) As UMR_Action
    Dim tempaction As UMR_Action, temp As Long
    If Len(Profile) = 0 Then Profile = CurrentProfile
    tempaction.ActionType = UMR_DoNothing
    If UMR_ButtonExists(Button, Profile) Then
        Select Case LCase(UMR_ButtonProperty(Button, "ActionType", , Profile))
            Case "keyboard"
                tempaction.Button = vKeyString2ID(UMR_ButtonProperty(Button, "Button", , Profile))
                If tempaction.Button > -1 Then
                    tempaction.ActionType = UMR_Keyboard
                    tempaction.CTRL = ToBool(UMR_ButtonProperty(Button, "CTRL", , Profile))
                    tempaction.ALT = ToBool(UMR_ButtonProperty(Button, "ALT", , Profile))
                End If
            Case "mediakey"
                tempaction.Button = vKeyString2ID(UMR_ButtonProperty(Button, "Button", , Profile))
                If tempaction.Button > -1 Then tempaction.ActionType = UMR_MediaKey
            Case "mouse"
                tempaction.ActionType = UMR_Mouse
                Select Case LCase(UMR_ButtonProperty(Button, "Button", , Profile))
                    Case "scroll up":     tempaction.Button = UMR_ScrollUp ' Mouse_Scroll -1
                    Case "scroll left":   tempaction.Button = UMR_ScrollLeft 'Mouse_Scroll -1, True
                    Case "scroll right":  tempaction.Button = UMR_ScrollRight 'Mouse_Scroll 1, True
                    Case "scroll down":   tempaction.Button = UMR_ScrollDown 'Mouse_Scroll 1
                    Case "click left":    tempaction.Button = UMR_LeftButton 'Mouse_Click vbLeftButton
                    Case "click middle":  tempaction.Button = UMR_MiddleButton 'Mouse_Click vbMiddleButton
                    Case "click right":   tempaction.Button = UMR_RightButton 'Mouse_Click vbRightButton
                    Case "move up":       tempaction.Button = UMR_MoveUp 'Mouse_MoveTo 0, -5, True
                    Case "move down":     tempaction.Button = UMR_MoveLeft 'Mouse_MoveTo 0, 5, True
                    Case "move left":     tempaction.Button = UMR_MoveRight 'Mouse_MoveTo -5, 0, True
                    Case "move right":    tempaction.Button = UMR_MoveDown 'Mouse_MoveTo 5, 0, True
                    Case Else: tempaction.ActionType = UMR_DoNothing
                End Select
        End Select
    End If
    UMR_CacheButton = tempaction
End Function
Public Function UMR_CacheButtons(Optional Profile As String)
    Dim tempstr() As String, temp As Integer
    If Len(Profile) = 0 Then Profile = CurrentProfile
    tempstr = Split("DPAD UP,DPAD LEFT,DPAD RIGHT,DPAD DOWN,CROSS,CIRCLE,SQUARE,TRIANGLE,L SHOULDER,R SHOULDER,START,SELECT,MENU,VOL UP,VOL DOWN", ",")
    For temp = 0 To UBound(tempstr)
        UMR_ButtonCache(temp) = UMR_CacheButton(tempstr(temp), Profile)
    Next
End Function
Public Function UMR_PerformCachedAction(Action As UMR_Action, isDown As Boolean, WasDown As Boolean)
    Select Case Action.ActionType
        Case UMR_DoNothing
        Case UMR_Keyboard, UMR_MediaKey
            VKeyPress Action.Button, Action.CTRL, Action.ALT, -1, isDown, Not isDown, WasDown
        Case UMR_Mouse
            Select Case Action.Button
                Case UMR_ScrollUp:      If isDown Then Mouse_Scroll -1
                Case UMR_ScrollLeft:    If isDown Then Mouse_Scroll -1, True
                Case UMR_ScrollRight:   If isDown Then Mouse_Scroll 1, True
                Case UMR_ScrollDown:    If isDown Then Mouse_Scroll 1
                
                Case UMR_MoveUp:        If isDown Then Mouse_MoveTo 0, -5, False
                Case UMR_MoveLeft:      If isDown Then Mouse_MoveTo -5, 0, False
                Case UMR_MoveRight:     If isDown Then Mouse_MoveTo 5, 0, False
                Case UMR_MoveDown:      If isDown Then Mouse_MoveTo 0, 5, False
                
                Case UMR_LeftButton:    UMR_DoMouse vbLeftButton, isDown, WasDown
                Case UMR_MiddleButton:  UMR_DoMouse vbMiddleButton, isDown, WasDown
                Case UMR_RightButton:   UMR_DoMouse vbRightButton, isDown, WasDown
            End Select
    End Select
End Function
Public Function UMR_DoMouse(Button As MouseButtonConstants, isDown As Boolean, WasDown As Boolean)
    If isDown Then
        If Not WasDown Then Mouse_Click Button, True, False: Debug.Print "CLICKING " & Button
    Else
        Mouse_Click Button, False, True: Debug.Print "RELEASING " & Button
    End If
End Function
Public Function UMR_HandleGamemode(State As String) As String
    Dim temp As Long, temp2 As Long, Nstate As String, Ostate As String, Dstate As String, tempstr As String     '0 is not pushed, 1 is down, 2 is released
    If Len(State) = 0 Then 'release all buttons currently down
        temp2 = Len(OldState)
        For temp = 1 To temp2
            If Mid(OldState, temp, 1) = 1 Then 'was down, release it
                UMR_PerformCachedAction UMR_ButtonCache(temp), False, True
            End If
        Next
    Else
        For temp = 1 To Len(State)
            Nstate = Mid(State, temp, 1)
            Dstate = Nstate
            If Nstate = "2" Then 'is not pressed and was pressed(released)
                UMR_PerformCachedAction UMR_ButtonCache(temp), False, True
                Dstate = "0"
            ElseIf Len(OldState) = 0 Then 'is pressed and was not pressed
                If Nstate = "1" Then UMR_PerformCachedAction UMR_ButtonCache(temp), True, False
            ElseIf Nstate = "1" Then
                Ostate = Mid(OldState, temp, 1)
                UMR_PerformCachedAction UMR_ButtonCache(temp - 1), True, Ostate = "1"
            End If
            tempstr = tempstr & Dstate
        Next
    End If
    'Debug.Print State
    OldState = State 'tempstr
    UMR_HandleGamemode = tempstr
End Function












Public Function UMR_GetSetting(ByVal Section As String, Key As String, Optional Default As String, Optional Save As Boolean) As String
    If Save Then
        If Len(Section) = 0 Then Section = "Settings" Else Section = "Settings\" & Section
        UMHini.SetKey Section, Key, Default
        UMR_GetSetting = Default
    Else
        UMR_GetSetting = UMHini.GetKey("Settings\" & Section, Key, Default)
    End If
End Function

Public Function CurrentWindow(IgnoreHwnd As Long) As String
    Static OldWindow As Long, OldLabel As String
    Dim temp As Long
    temp = GetForegroundWindow
    IsInFocus = (temp = IgnoreHwnd)
    If temp <> IgnoreHwnd Then
        OldWindow = temp
        OldLabel = GetHwndEXE(temp)
        OldLabel = Right(OldLabel, Len(OldLabel) - InStrRev(OldLabel, "\"))
    End If
    CurrentWindow = OldLabel
End Function

Public Sub UMR_EnumPrograms(Optional ListId As Long = 3)
    Dim temp As Long, tempstr As String
    LCAR_ClearList ListId
    fEnumWindows True
    For temp = 0 To WindowCount - 1
        With WindowList(temp)
            tempstr = GetHwndEXE(.hwnd)
            tempstr = Right(tempstr, Len(tempstr) - InStrRev(tempstr, "\"))
            If LCAR_FindListItemByName(ListId, tempstr) = -1 Then LCAR_AddListItem ListId, tempstr, , , , CStr(temp)
        End With
    Next
    HideAllLists ListId
End Sub

Public Sub UMR_LoadHINI(CTL As Hini, Optional ListId As Long = 1)
    Set UMHini = CTL
    If UMHini.LoadFile(App.Path & "\settings.hini") Then 'Debug.Print "Load failed"
        If ListId > -1 Then UMR_EnumProfiles ListId
    Else
        UMHini.CreateSection "Profiles"
    End If
End Sub

Public Sub UMR_SaveHINI()
    UMHini.SaveFile App.Path & "\settings.hini"
End Sub

Public Function UMR_EnumProfiles(Optional ListId As Long = 1)
    Dim temp As Long, Count As Long, List() As String
    LCAR_ClearList ListId
    Count = UMHini.EnumSections("Profiles", List)
    If Not UMR_ProfileExists("Default") Then LCAR_AddListItem ListId, "Default"
    For temp = 1 To Count
        LCAR_AddListItem ListId, List(temp)
    Next
End Function

Public Function UMR_NewProfile(Profile As String, Optional ListId As Long = 1) As Boolean
    If Not UMHini.SectionExists("Profiles\" & Profile) Then
        UMHini.CreateSection "Profiles\" & Profile
        UMR_NewProfile = True
        If ListId > -1 Then LCAR_AddListItem ListId, Profile
    End If
    UMR_LoadProfile Profile
End Function

Public Function UMR_DeleteProfile(Optional Profile As String, Optional ListId As Long = 1) As Boolean
    Dim temp As Long
    If Len(Profile) = 0 Then
        Profile = CurrentProfile
        UMR_LoadProfile Empty
    End If
    If UMHini.SectionExists("Profiles\" & Profile) Then
        UMHini.DeleteSection "Profiles\" & Profile
        If ListId > -1 Then
            If StrComp(Profile, "Default", vbTextCompare) <> 0 Then
                temp = LCAR_FindListItemByName(ListId, Profile)
                If temp > -1 Then LCAR_DeleteListItem ListId, temp
            End If
        End If
    End If
End Function

Public Sub UMR_LoadProfile(Optional Profile As String, Optional Sock As CSocketMaster)
    Dim temp As Long, ButtonList() As String, ButtonCount As Long, doit As Boolean
    If Sock Is Nothing Then
        doit = True
    Else
        If Sock.IsTheServer Then doit = True 'CurrentProfile = Profile
    End If
    
    If doit Then
        HideAllLists 0
        LCAR_ClearList 0
        CurrentProfile = Profile
        LCAR_SetText "btnpath", "Current profile: " & Profile, 3, 2
    End If

    If Len(Profile) > 0 Then
        If Not UMR_ProfileExists(Profile) Then Profile = "Default"
    
        ButtonCount = UMHini.EnumSections("Profiles\" & Profile, ButtonList, "ColorValue")
        For temp = 1 To ButtonCount
            If Sock Is Nothing Then
                LCAR_AddListItem 0, ButtonList(1, temp), , Val(ButtonList(2, temp))
            Else
                SendData Sock, "addbutton """ & ButtonList(1, temp) & """ " & Val(ButtonList(2, temp))
            End If
        Next
        
        UMR_CacheButtons Profile
    End If
End Sub

Sub SendData(Sock As CSocketMaster, Data As String)
    Sock.SendData Data & Chr(10)
End Sub

Public Sub UMR_SaveButton(Profile As String, Button As String, ColorValue As Long, Optional Color As String, Optional ActionType As String, Optional CTRL As String, Optional ALT As String, Optional ButtonID As String)
    UMHini.SetKey "Profiles\" & Profile, "Stardate", CStr(StarDate(Now, 5))
    UMHini.SetKey "Profiles\" & Profile & "\" & Button, "ColorValue", CStr(ColorValue)
    If Len(Color) = 0 Then
        LCAR_AddListItem 0, Button, , ColorValue
        UMHini.SetKey "Profiles\" & Profile & "\" & Button, "MissingData", "True"
    Else
        UMHini.SetKey "Profiles\" & Profile & "\" & Button, "MissingData", "False"
        UMHini.SetKey "Profiles\" & Profile & "\" & Button, "Color", Color
        UMHini.SetKey "Profiles\" & Profile & "\" & Button, "ActionType", ActionType
        UMHini.SetKey "Profiles\" & Profile & "\" & Button, "CTRL", CTRL
        UMHini.SetKey "Profiles\" & Profile & "\" & Button, "ALT", ALT
        UMHini.SetKey "Profiles\" & Profile & "\" & Button, "Button", ButtonID
    End If
End Sub

Public Function UMR_ButtonData(Profile As String, Button As String, Optional Sock As CSocketMaster) As String
    Dim tempstr As String
    If Not Sock Is Nothing Then tempstr = "requestbutton " & Profile & " " & Button & " "
    tempstr = tempstr & UMR_ButtonProperty(Button, "ColorValue", , Profile) & " "
    tempstr = tempstr & UMR_ButtonProperty(Button, "Color", , Profile) & " "
    tempstr = tempstr & UMR_ButtonProperty(Button, "ActionType", , Profile) & " "
    tempstr = tempstr & UMR_ButtonProperty(Button, "CTRL", , Profile) & " "
    tempstr = tempstr & UMR_ButtonProperty(Button, "ALT", , Profile) & " "
    tempstr = tempstr & UMR_ButtonProperty(Button, "Button", , Profile)
    If Not Sock Is Nothing Then SendData Sock, tempstr
    UMR_ButtonData = tempstr
End Function

Public Function UMR_ButtonProperty(Button As String, Property As String, Optional Default As String, Optional Profile As String, Optional Save As Boolean) As String
    If Len(Profile) = 0 Then Profile = CurrentProfile
    If Len(Profile) = 0 Then Profile = "Default"
    If Save Then
        UMHini.SetKey "Profiles\" & Profile, "Stardate", CStr(StarDate(Now, 5))
        UMHini.SetKey "Profiles\" & Profile & "\" & Button, Property, Default
    Else
        If Len(Button) = 0 Then
            UMR_ButtonProperty = Default
        Else
            UMR_ButtonProperty = UMHini.GetKey("Profiles\" & Profile & "\" & Button, Property, Default)
        End If
    End If
End Function

Public Function UMR_StarDate(Profile As String) As Double
    UMR_StarDate = CDbl(UMHini.GetKey("Profiles\" & Profile, "Stardate", "-99999999"))
End Function

Public Function UMR_ButtonExists(Button As String, Optional Profile As String)
    If Len(Profile) = 0 Then Profile = CurrentProfile
    UMR_ButtonExists = UMHini.SectionExists("Profiles\" & Profile & "\" & Button)
End Function

Public Sub UMR_DeleteButton(Button As String, Optional Profile As String)
    Dim temp As Long
    If Len(Profile) = 0 Then Profile = CurrentProfile
    UMHini.DeleteSection "Profiles\" & Profile & "\" & Button
    If StrComp(Profile, CurrentProfile, vbTextCompare) = 0 Then
        temp = LCAR_FindListItemByName(0, Button)
        If temp > -1 Then LCAR_DeleteListItem 0, temp
    End If
End Sub

Public Function ToBool(Text As String) As Boolean
    If StrComp(Text, "yes", vbTextCompare) = 0 Then ToBool = True
End Function

Public Function UMR_ProfileExists(Profile As String) As Boolean
    UMR_ProfileExists = UMHini.SectionExists("Profiles\" & Profile)
End Function

Public Sub UMR_Execute(Button As String, Optional Profile As String)
    Dim temp As Long
    If Len(Profile) = 0 Then Profile = CurrentProfile
    
    Select Case LCase(UMR_ButtonProperty(Button, "ActionType", , Profile))
        Case "keyboard"
            temp = vKeyString2ID(UMR_ButtonProperty(Button, "Button", , Profile))
            If temp > -1 Then VKeyPress temp, ToBool(UMR_ButtonProperty(Button, "CTRL", , Profile)), ToBool(UMR_ButtonProperty(Button, "ALT", , Profile))
        Case "mediakey"
            temp = vKeyString2ID(UMR_ButtonProperty(Button, "Button", , Profile))
            If temp > -1 Then VKeyPress temp
        Case "mouse"
            Select Case LCase(UMR_ButtonProperty(Button, "Button", , Profile))
                Case "scroll up":     Mouse_Scroll -1
                Case "scroll left":   Mouse_Scroll -1, True
                Case "scroll right":  Mouse_Scroll 1, True
                Case "scroll down":   Mouse_Scroll 1
                Case "click left":    Mouse_Click vbLeftButton
                Case "click middle":  Mouse_Click vbMiddleButton
                Case "click right":   Mouse_Click vbRightButton
                Case "move up":       Mouse_MoveTo 0, -5, False
                Case "move down":     Mouse_MoveTo 0, 5, False
                Case "move left":     Mouse_MoveTo -5, 0, False
                Case "move right":    Mouse_MoveTo 5, 0, False
            End Select
    End Select
End Sub

Public Function UMR_SendTasks(Optional Sock As CSocketMaster, Optional ByVal HTML As String) As String
    Dim temp As Long, Count As Long, tempstr As String
    Count = fEnumWindows(True)
    For temp = 0 To Count - 1
        If Len(HTML) = 0 Then
            SendData Sock, "task " & WindowList(temp).hwnd & " '" & WindowList(temp).WindowName & "' 1"
        Else
            tempstr = tempstr & UMR_GenerateHTML(HTML, "Command", "?switchto%20" & CStr(WindowList(temp).hwnd), "WindowName", WindowList(temp).WindowName)
        End If
    Next
    UMR_SendTasks = tempstr
End Function

Public Function UMR_GenerateHTML(ByVal HTML As String, ParamArray Varlist() As Variant) As String
    Dim temp As Long
    For temp = 0 To UBound(Varlist) Step 2
        If temp < UBound(Varlist) Then HTML = Replace(HTML, "%" & Varlist(temp) & "%", Varlist(temp + 1), , , vbTextCompare)
    Next
    HTML = Replace(HTML, "%t%", Chr(9), , , vbTextCompare) 'tab
    UMR_GenerateHTML = Replace(HTML, "%n%", vbNewLine, , , vbTextCompare) 'newline
End Function


Public Function IsADir(Filename As String) As Boolean
    On Error Resume Next
    If Len(Filename) > 0 Then IsADir = (GetAttr(Filename) And vbDirectory) = vbDirectory
End Function
Public Function UMR_SendFiles(Optional Sock As CSocketMaster, Optional Path As String, Optional DRV As DriveListBox, Optional DIR As DirListBox, Optional File As FileListBox, Optional DirHTML As String, Optional HTML As String, Optional Start As Long, Optional Amount As Long = -1, Optional ByRef tempstr As String) As Boolean
    Dim temp As Long, Count As Long
    Static LastBrowsed As String
    On Error Resume Next
    
    If Len(DirHTML) = 0 Then DirHTML = HTML
    
    If Len(Path) = 0 Then
        DRV.Refresh
        For temp = 0 To DRV.ListCount - 1
            tempstr = tempstr & UMR_SendFile(Sock, Left(DRV.List(temp), 2), True, DirHTML)
        Next
        tempstr = tempstr & UMR_SendFile(Sock, ShellFolder("Desktop"), True, DirHTML)
        tempstr = tempstr & UMR_SendFile(Sock, ShellFolder, True, HTML)
        tempstr = tempstr & UMR_SendFile(Sock, ShellFolder("My Music"), True, DirHTML)
        tempstr = tempstr & UMR_SendFile(Sock, ShellFolder("My Pictures"), True, DirHTML)
        tempstr = tempstr & UMR_SendFile(Sock, ShellFolder("My Video"), True, DirHTML)
        If Len(LastBrowsed) > 0 Then tempstr = tempstr & UMR_SendFile(Sock, LastBrowsed, True, DirHTML)
    
    
    ElseIf Not IsADir(Path) Then
        UMR_SendFiles = True
        Exit Function
    Else

    If StrComp(Path, DIR.Path, vbTextCompare) = 0 Then
        DIR.Refresh
        File.Refresh
    Else
        If Len(Path) = 2 And Right(Path, 1) = ":" Then Path = Path & "\"
        DIR.Path = Path
        File.Path = Path
    End If
    
    If Len(Path) > 3 Then tempstr = tempstr & UMR_SendFile(Sock, Left(Path, InStrRev(Path, "\") - 1), True, DirHTML, 0)
    LastBrowsed = Path
    
    For temp = 0 To DIR.ListCount - 1
        If temp >= Start Then
            tempstr = tempstr & UMR_SendFile(Sock, DIR.List(temp), True, DirHTML)
            Count = Count + 1
        End If
        If Count > Amount And Amount > -1 Then Exit For
    Next
    
    For temp = 0 To File.ListCount - 1
        If Count > Amount And Amount > -1 Then
            Exit For
        ElseIf temp >= Start Then
            tempstr = tempstr & UMR_SendFile(Sock, Replace(Path & "\" & File.List(temp), "\\", "\"), False, HTML)
            Count = Count + 1
        End If
    Next
    
    End If
End Function
Public Function UMR_SendFile(Optional Sock As CSocketMaster, Optional Filename As String, Optional isDIR As Boolean, Optional ByVal HTML As String, Optional Number As Long = 1, Optional RootTitle As String = "..") As String
    Dim Title As String, fType As String, Size As Long, FILEsize As String, LCARSize As Long, Command As String
    If Len(Filename) > 0 Then
        If Len(HTML) = 0 Then
            If isDIR Then
                HTML = "folder dir '" & Filename & "' " & Number
            Else
                HTML = "folder " & FileLen(Filename) & " '" & Filename & "' " & Number
            End If
            SendData Sock, HTML
        Else
            'generate HTML
            Title = FileTitle(Filename, True)
            If Number = 0 Then Title = RootTitle
            If isDIR Then
                fType = "File Folder"
            Else
                fType = FileTypeName(FileExtention(Filename))
                Size = FileLen(Filename)
                LCARSize = SizeToLCAR(Size)
                FILEsize = SizeToText(Size, " Q", " K", "M", " G")
            End If
            Command = "?setdir " & Filename
            
            UMR_SendFile = UMR_GenerateHTML(HTML, "Command", Command, "Bytes", Size, "FileSize", FILEsize, "LCARsize", LCARSize, "Filename", Filename, "FileTitle", Title, "UTitle", UCase(Title), "Extention", FileExtention(Filename), "Type", fType)
        End If
        'Debug.Print HTML
    End If
End Function




Public Function UMR_ProcessTemplate(ByVal Filename As String, Optional DRV As DriveListBox, Optional DIR As DirListBox, Optional File As FileListBox, Optional Src As String) As String
    Dim HTML As String, Count As Long, tempstr() As String, temp As Long
    Dim DirHTML As String, FileHTML As String, RootName As String
    
    If InStr(Filename, ":\") = 0 Then Filename = Replace(App.Path & "\" & Filename, "\\", "\")
    Filename = Replace(Filename, "%profile%", IIf(Len(CurrentProfile) = 0, "Default", CurrentProfile), , , vbTextCompare)
    If StrComp(Filename, Src, vbTextCompare) = 0 Then Exit Function
    
    HTML = LoadFile(Filename)
    Count = QTAG_Split(HTML, tempstr)
    For temp = 0 To Count - 1
        'Debug.Print "Line " & temp & " " & QTAG_isTag(tempstr(temp)) & tempstr(temp)
        If QTAG_isTag(tempstr(temp)) Then
            Select Case LCase(QTAG_Name(tempstr(temp)))
                Case "stardate":    tempstr(temp) = StarDate(Now)
                Case "time":        tempstr(temp) = time
                Case "date":        tempstr(temp) = Format(Date, "dd/m/yyyy")
                Case "longdate":    tempstr(temp) = Format(Date, "dddd mmm dd, yyyy")
                Case "dir":         tempstr(temp) = CurrentDir
                Case "title":       tempstr(temp) = FileTitle(CurrentDir)
                Case "version":     tempstr(temp) = App.Major & "." & App.Minor & "." & App.Revision
                Case "about":       tempstr(temp) = "UMRemote is programmed by Techni Myoko"
                Case "mysite":      tempstr(temp) = "http://sites.google.com/site/neotechni/"
                Case "profile":     tempstr(temp) = IIf(Len(CurrentProfile) = 0, "Default", CurrentProfile)
                
                Case "include"
                    FileHTML = QTAG_GetValue(tempstr(temp), "src")
                    If Len(FileHTML) > 0 Then FileHTML = UMR_ProcessTemplate(FileHTML, DRV, DIR, File, Filename)
                    If Len(FileHTML) > 0 Then tempstr(temp) = FileHTML
                
                Case "buttonlist" 'accepted subtags: Command, Name, Color
                    FileHTML = QTAG_GetValue(tempstr(temp), "HTML", HTML_Butnlist)
                    tempstr(temp) = UMR_SendProfile(, FileHTML)
                
                Case "tasklist" 'accepted subtags: Command, WindowName
                    FileHTML = QTAG_GetValue(tempstr(temp), "HTML", HTML_Tasklist)
                    tempstr(temp) = UMR_SendTasks(, FileHTML)
                    
                Case "filelist" 'accepted subtags: Command, Bytes, Filesize, LCARsize, FileTitle, UTitle, Extention, Type
                    FileHTML = QTAG_GetValue(tempstr(temp), "HTML", HTML_Filelist)
                    DirHTML = QTAG_GetValue(tempstr(temp), "dirHTML", FileHTML)
                    RootName = QTAG_GetValue(tempstr(temp), "RootName", "..")
                    UMR_SendFiles , CurrentDir, DRV, DIR, File, DirHTML, FileHTML, , , tempstr(temp)
                
            End Select
            'accepted subtags: n (newline), t (tab)
        End If
    Next
    UMR_ProcessTemplate = Join(tempstr, Empty)
End Function

Public Function UMR_SendProfile(Optional Profile As String, Optional ByVal HTML As String = HTML_Butnlist) As String
    Dim temp As Long, ButtonList() As String, ButtonCount As Long, tempstr As String, Command As String, Name As String, Color As String
    If Len(Profile) = 0 Then
        Profile = CurrentProfile
        If Len(Profile) = 0 Then Profile = "default"
    End If
    ButtonCount = UMHini.EnumSections("Profiles\" & Profile, ButtonList, "ColorValue")
    For temp = 1 To ButtonCount
        Name = ButtonList(1, temp)
        Color = HTMLColor(Val(ButtonList(2, temp)))
        Command = "exe " & Name
        tempstr = tempstr & UMR_GenerateHTML(HTML, "Command", Command, "Name", Name, "Color", Color)
    Next
    UMR_SendProfile = tempstr
End Function
