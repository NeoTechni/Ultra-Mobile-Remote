VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Ultra Mobile Remote"
   ClientHeight    =   1185
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2055
   FillStyle       =   0  'Solid
   BeginProperty Font 
      Name            =   "LCARS"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   79
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   137
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.FileListBox Filmain 
      Height          =   225
      Left            =   960
      TabIndex        =   3
      Top             =   840
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.DirListBox Dirmain 
      Height          =   345
      Left            =   480
      TabIndex        =   2
      Top             =   720
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.DriveListBox drvmain 
      Height          =   345
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Visible         =   0   'False
      Width           =   270
   End
   Begin Project1.Hini Hinimain 
      Left            =   1440
      Top             =   120
      _ExtentX        =   847
      _ExtentY        =   847
   End
   Begin VB.PictureBox picbuffer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   1560
      ScaleHeight     =   17
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   25
      TabIndex        =   0
      Top             =   360
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.Timer TimerEffects 
      Interval        =   1
      Left            =   480
      Top             =   120
   End
   Begin VB.Timer TimerBlink 
      Interval        =   250
      Left            =   120
      Top             =   120
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'?curdir("D:\")

Dim WithEvents RC As SpSharedRecoContext
Attribute RC.VB_VarHelpID = -1
Dim myGrammar As ISpeechRecoGrammar

Public WithEvents Sock As CSocketMaster, OldWindow As String
Attribute Sock.VB_VarHelpID = -1

Dim wasrotated As Boolean, wasminimized As Boolean, GameMode As String, rawmode As Boolean
Dim oldCol As Long, oldRow As Long, LockedOn As Boolean, CPath As String, LastSelected As Long

Public LCARname As String, LCARindex As Long, LCARid As Long, LCARGroup As Long, LCARblinking As Boolean, LCARitem As Long
Public Sound As Long, OldButton As String, ValidDisconnect As Boolean, IPAddress As String, Port As Long, Client As Boolean

Public Event LCARClicked(Name As String, Index As Long)
Public Event LCARMouseDown(Name As String, Index As Long)
Public Event LCARMouseUp(Name As String, Index As Long)
Public Event LCARMouseScroll(Name As String, Index As Long)

    Const LCAR_SMB As Long = LCAR_Purple
    Const LCAR_CMD As Long = LCAR_DarkPurple
    Const LCAR_ABC As Long = LCAR_Orange
    Const LCAR_123 As Long = LCAR_LightBlue

Sub SetupVoiceRecognition()
    Set RC = New SpSharedRecoContext
    Set myGrammar = RC.CreateGrammar
    myGrammar.DictationSetState SGDSActive
End Sub


Private Sub RC_Recognition(ByVal StreamNumber As Long, ByVal StreamPosition As Variant, ByVal RecognitionType As SpeechLib.SpeechRecognitionType, ByVal Result As SpeechLib.ISpeechRecoResult)
    Dim tempstr As String
    tempstr = Result.PhraseInfo.GetText
    SendData "umr " & LCase(tempstr)
End Sub



Sub HandleGameMode()
    Dim Before As String
    Before = GameMode
    'UMR_Execute GetParam(tempstr, 2, CurrentProfile), tempstr(1)
    '"DPAD UP", "DPAD LEFT", "DPAD RIGHT", "DPAD DOWN", "CROSS", "CIRCLE", "SQUARE", "TRIANGLE", "L SHOULDER", "R SHOULDER", "START", "SELECT", "MENU", "VOL UP", "VOL DOWN"
    GameMode = UMR_HandleGamemode(GameMode)
    'If Len(Before) > 0 Then
    '    If Before <> "000000000000000" Then Debug.Print "Before: " & Before & " After: " & GameMode
    'End If
    If StrComp(Before, GameMode, vbBinaryCompare) <> 0 Then UMR_HandleGamemode GameMode
End Sub



Private Sub ShowConnectDialog()
    HideAllLists 9
    LCARlists(10).Visible = True
End Sub

Sub MakeButton(Name As String)
    If Not UMR_ButtonExists(Name) Then
        UMR_SaveButton CurrentProfile, Name, LCAR_Orange
    End If
End Sub
Sub MakeButtons(ParamArray Buttons() As Variant)
    Dim temp As Long
    For temp = 0 To UBound(Buttons)
        MakeButton (Buttons(temp))
    Next
    UMR_LoadProfile CurrentProfile
End Sub

Public Sub LCARClicked()
    Static docut As Boolean
    Const NoProfile As String = "No profile has been selected"
    Dim RAND As Boolean, temp As Long, Path As String, tempstr As String, Caption As String
    Caption = LCAR_ButtonList(LCARid).Text
    RAND = True
    Sound = -1
    Select Case LCase(LCARname)
        Case "btnsearch"
            SetRedAlert False
            
        Case "btnpath" ' Directory path
            Select Case LCARindex
                Case 0 'new profile
                    HideKeyboard
                    UMR_EnumPrograms
                Case 1 'load profile
                    HideKeyboard
                    HideAllLists 1
                Case 2 'delete profile
                    If Len(CurrentProfile) > 0 Then
                        Prompt "This will cause damage to the profile system", True, True, , "deleteprofile"
                    Else
                        Operation = Empty
                        Prompt NoProfile, False, True
                    End If
                Case 3 'current profile
                    Select Case Sock.State
                        Case sckClosed: tempstr = "Closed"
                        Case sckClosing: tempstr = "Closing"
                        Case sckConnected: tempstr = "Connected"
                        Case sckConnecting: tempstr = "Connecting"
                        Case sckConnectionPending: tempstr = "Pending"
                        Case sckError: tempstr = "Error"
                        Case sckHostResolved: tempstr = "Resolved"
                        Case sckListening: tempstr = "Listening"
                        Case sckOpen: tempstr = "Open"
                        Case sckResolvingHost: tempstr = "Resolving"
                    End Select
                    Prompt "This is your current connection state", False, False, tempstr
                Case 4 'IP address
                    Prompt "This is your local IP address and port: " & Sock.LocalPort, False, False, Sock.LocalIP
            End Select
            
        Case "btntasks" 'quicktasks
            Select Case LCARindex '> 0 And LCARindex < 10 And LCARindex <> 6 And LCARindex <> 2 And LCARlists(0).Visible Then
                Case 0
                    If Sock.IsTheServer Or Not Sock.IsTheClient Then
                        Prompt "This program cannot serve as a keyboard on the server", False, True, "PEBKAC"
                    Else
                        ShowKeyboard Empty, "send"
                    End If
                Case 1
                    If Sock.IsTheServer Or Not Sock.IsTheClient Then
                        Prompt "This program cannot serve as a numboard on the server", False, True, "PEBKAC"
                    Else
                        ShowKeyboard "0", "send", True
                    End If
                Case 2 'add button
                    If Len(CurrentProfile) = 0 Then
                        Prompt NoProfile, False, True
                    Else
                        OldButton = Empty
                        If isRotated Then RotateScreen
                        HideKeyboard
                        ResetButtonOptions
                        HideAllLists 7
                    End If
                Case 3 'edit button
                    If Len(CurrentProfile) = 0 Then
                        Prompt NoProfile, False, True
                    Else
                        SetRedAlert
                        Sound = 104
                        RefreshPreview "The next button you press will be edited"
                        Operation = "editbutton"
                    End If
                Case 4 'delete button
                     If Len(CurrentProfile) = 0 Then
                        Prompt NoProfile, False, True
                    Else
                        SetRedAlert
                        Sound = 104
                        RefreshPreview "The next button you press will be deleted"
                        Operation = "deletebutton"
                    End If
                Case 5 ' XPERIA PLAY
                    MakeButtons "DPAD UP", "DPAD LEFT", "DPAD RIGHT", "DPAD DOWN", "CROSS", "CIRCLE", "SQUARE", "TRIANGLE", "L SHOULDER", "R SHOULDER", "START", "SELECT", "MENU", "VOL UP", "VOL DOWN"
                Case 6 'task list
                    If Sock.IsTheServer Or Not Sock.IsTheClient Then
                        Prompt "See the taskbar down there? Use it", False, True, "PEBKAC"
                    Else
                        SendData "task"
                        LCAR_ClearList 11
                        HideAllLists 11
                    End If
                Case 7 'file list
                    If Sock.IsTheServer Or Not Sock.IsTheClient Then
                        Prompt "Press Windows+D, then double-click 'My Computer'", False, True, "PEBKAC"
                    Else
                        SendData "folder dir"
                        HideAllLists 12
                        LCAR_ClearList 12
                    End If
            End Select
            
            
        Case "btnmenu" 'menu bar
            Select Case LCARindex
                Case 0 'Connect to server
                    ShowConnectDialog
                    
                Case 1 'options
                    If isRotated Then RotateScreen
                    HideKeyboard
                    HideAllLists 4
                Case 2 'about
                    Prompt "UMRemote is programmed by Techni Myoko", False, False, App.Major & "." & App.Minor & "." & App.Revision
                Case 4: WindowState = vbMinimized 'Minimize
                Case 5: Unload Me 'Exit
            End Select
            
        Case "btnsystem"
            Select Case LCARindex
                Case 0 'Toggle multiselect
                    LCARlists(0).MultiSelect = Not LCARlists(0).MultiSelect
                    LCAR_ButtonList(LCARid).State = IIf(LCAR_ButtonList(LCARid).State = -1, 0, -1)
                    If Not LCARlists(0).MultiSelect And LCARlists(0).SelectedItems > 1 Then
                        LCAR_Select 0, 0
                        If LCARlists(0).SelectedItem > -1 Then LCAR_SelectItem 0, LCARlists(0).SelectedItem
                    End If
                Case 1: LCAR_Select 0, 0 'Select none
                Case 2 'Select All
                    LCARlists(0).MultiSelect = True
                    LCAR_Select 0, 1
                    LCAR_ButtonList(LCAR_FindLCAR("btnsystem")).State = -1
                Case 3 'Invert selected
                    LCARlists(0).MultiSelect = True
                    LCAR_Select 0, -1
                    LCAR_ButtonList(LCAR_FindLCAR("btnsystem")).State = -1
                Case 4 'TestCirLCAR

                Case 5
                    'ResizeLCARs
                    RotateScreen
            End Select
            RefreshPreview
        
        
        
         Case "frmnumbottom" 'numboard
            Select Case LCARindex
                Case 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 11, 12, 13, 14, 15: ProcessKey Caption
                Case 10 'ok
                    If ValidateNumber Then
                        HideKeyboard
                        Select Case LCase(Operation)
                            Case Empty, "send"
                            Case "editip"
                                LCARlists(9).ListItems(LCARlists(9).SelectedItem).FileLCAR = Text
                                UMR_GetSetting "Defaults", "IP" & LCARlists(9).SelectedItem, Text, True
                                LCAR_Select 10, 0
                                ShowConnectDialog
                            Case "editport"
                                LCARlists(10).ListItems(0).FileLCAR = Text
                                UMR_GetSetting "Defaults", "Port", Text, True
                                LCAR_Select 9, 0
                                ShowConnectDialog
                            Case "setport"
                                Sock.AutoBind Val(Text), True, True
                                UMR_GetSetting "Defaults", "LocalPort", Text, True
                                LCARlists(4).ListItems(4).Side = Text
                                HideAllLists 4
                            
                            Case Else: MsgBox OldText & vbNewLine & Operation & vbNewLine & Text
                        End Select
                        
                    End If
                    Select Case LCase(Operation)
                        Case "send"
                            If Sock.IsTheClient Then SendData "sendtext '" & Text & "' 1"
                    End Select
                Case Else: Debug.Print LCARindex
            End Select
        
        Case "frmbottom" 'keyboard
            
            Select Case LCARindex
                Case 0, 1, 2, 3, 6, 7, 8: ProcessKey Caption
                Case 11: ProcessKey " " 'Spacebar
               
                Case 4 'Symbols
                    LCAR_Blink LCARid, Not LCAR_isBlinking(LCARid)
                    Symbols = LCAR_isBlinking(LCARid)
                Case 5 'caps
                    LCAR_Blink LCARid, Not LCAR_isBlinking(LCARid)
                    Caps = LCAR_isBlinking(LCARid)
                    
                Case 9 'cancel
                    ProcessKey Caption
                    Select Case LCase(Operation)
                        Case "namebutton"
                            HideAllLists 7
                    End Select
                Case 10 'OK
                    HideKeyboard
                    Select Case LCase(Operation)
                        Case Empty
                        Case "namebutton"
                            LCARlists(7).ListItems(0).Side = Text
                            HideAllLists 7
                        Case "send"
                            If Sock.IsTheClient Then SendData "sendtext '" & Text & "' 1"
                        Case Else: MsgBox OldText & vbNewLine & Operation & vbNewLine & Text
                    End Select
            End Select
            
            
           
            
        Case "btnaprompt"
            If LCARindex = 1 Then
                Select Case LCase(Operation)
                    Case "deleteprofile": UMR_DeleteProfile
                    Case "deletebutton": UMR_DeleteButton OldButton
                    Case "reconnect": Sock.Connect IPAddress, Port
                End Select
            Else

            End If
            HideGroup 7
            HideAllLists 0
            SetRedAlert False


        Case "btnmouse"
            If LCARindex = 3 Then
                Scrolling = Not Scrolling
                LCAR_Blink LCARid, Scrolling
            Else
                LCAR_Blink LCAR_FindLCAR("btnmouse", 9, 0), LCARindex = 0
                LCAR_Blink LCAR_FindLCAR("btnmouse", 9, 1), LCARindex = 1
                LCAR_Blink LCAR_FindLCAR("btnmouse", 9, 2), LCARindex = 2
                Select Case LCARindex
                    Case 0: Button = "left"
                    Case 1: Button = "middle"
                    Case 2: Button = "right"
                End Select
            End If
    End Select
    
    If Sound > -1 Then RAND = False
    If RAND Then PlayRandomSound 101, 102, 103, 106, 107, 108, 109 Else PlayRandomSound Sound
End Sub

Public Sub LCARListItemClicked()
    'Debug.Print ClickedSide
    Dim Filename As String, Size As Long, Item As Long, RAND As Boolean, Caption As String, LCARtext As String
    Item = LCARlists(ListId).SelectedItem
    If Item = -1 Then Item = LCARitem

    With LCARlists(ListId).ListItems(Item)
        Filename = .Tag
        Size = .Size
        Caption = .Text
        LCARtext = CStr(.FileLCAR)
    End With
    
    
    RAND = True
    Select Case ListId
        Case 0 'button List
            OldButton = Caption
            Select Case LCase(Operation)
                Case "editbutton"
                    SetRedAlert False
                    ResetButtonOptions Caption
                    If isRotated Then RotateScreen
                    HideAllLists 7
                    Operation = Empty
                Case "deletebutton"
                    Prompt "This will cause damage to the button system", True, True
                Case Else
                    'MsgBox "EXECUTING THIS BUTTON! " & Caption
                    'UMR_Execute Caption
                    Operation = Empty
                    SendData "execute """ & CurrentProfile & """ """ & Caption & """"
            End Select
            
            
        Case 1 'profile list
            UMR_LoadProfile Caption
            
        
        Case 2 'keyboard
            pX = (oldCol + 1) / 11
            pY = (oldRow + 1) / 5
            
            Filename = Text
            ProcessKey GetKey(Item)
            'MsgBox "Before: " & Filename & vbNewLine & "After: " & Text
            LCAR_SetText "btntxtbox", Text, , , , , , , , False
            
        Case 3 'new profile\running exe list
            UMR_NewProfile Caption
            
        Case 4 'Menu
            LCAR_ClearList 5
            Select Case Item
                Case 0 'Audio
                    LCAR_AddListItems 5, "On", "Off"
                    LCAR_SelectItem 5, Val(IIf(Mute, 1, 0))
                    RefreshPreview "This options enables/disables sounds"
                Case 1 'Filesizes
                    LCAR_AddListItems 5, "On", "Off"
                    LCAR_SelectItem 5, Val(IIf(SizeMode, 0, 1))
                    RefreshPreview "This makes filesizes either appear like the show or easy to read"
                Case 2 'font size
                    LCAR_AddListItems 5, "0", "1", "2"
                    LCAR_SelectItem 5, OldInc
                    RefreshPreview "This lets you increase/decrease the default font size"
                Case 3 'UI mode
                    LCAR_AddListItems 5, "Classic", "Classic+AA", "Nemesis"
                    If ThreeDmode Then
                        LCAR_SelectItem 5, 2
                    ElseIf AntiAliasing Then
                        LCAR_SelectItem 5, 1
                    End If
                    RefreshPreview "This lets you use antialiasing or the new 3D LCAR look"
                Case 4 'Port
                    ShowKeyboard CStr(Sock.LocalPort), "setport", True, 1, 999999
                Case 5 'tolerance
                    LCAR_AddListItems 5, "1", "5", "10", "15", "20", "25"
                    LCAR_SelectItem 5, LCAR_FindListItemByName(5, CStr(swTolerance))
                    RefreshPreview "How many pixels you must move before sending it to the server"
            End Select
            LCARlists(5).Visible = True
        
        Case 5 'menu option
            LCARlists(4).ListItems(LCARlists(4).SelectedItem).Side = LCARlists(5).ListItems(Item).Text
            Select Case LCARlists(4).SelectedItem
                Case 0 'Audio
                    Mute = Item = 1
                    SaveSetting "LCAR", "MAIN", "MUTE", CStr(Mute)
                Case 1 'Filesizes
                    SizeMode = Item = 0
                    SaveSetting "LCAR", "MAIN", "SizeMode", CStr(SizeMode)
                Case 2 'Font size
                    LCAR_FontIncrement Item
                    SaveSetting "LCAR", "MAIN", "FontSize", CStr(Item)
                Case 3 'UI mode
                    SetupUImode LCAR_SelectedItem(5)
                Case 5 'tolerance
                    swTolerance = Val(Caption)
            End Select
            
        Case 6 'numboard
            ProcessKey Filename
            
        Case 7 'button creator
            LCAR_ClearList 8
            Select Case Item
                Case 0 'Name
                    ShowKeyboard LCARlists(7).ListItems(0).Side, "namebutton"
                Case 1 'Color
                    AddColorsToList 8
                Case 2 'action type
                    LCAR_AddListItems 8, "Keyboard", "Mediakey", "Mouse"
                    LCARlists(7).ListItems(5).Side = "NONE"
                Case 3, 4 'CTRL & ALT
                    Select Case LCase(LCARlists(7).ListItems(2).Side)
                        Case "keyboard"
                            LCAR_AddListItems 8, "No", "Yes"
                    End Select
                Case 5 'button
                    Select Case LCase(LCARlists(7).ListItems(2).Side)
                        Case "keyboard"
                            AddRegdButtons 8
                        Case "mediakey"
                            AddRegdButtons 8, True
                        Case "mouse"
                            LCAR_AddListItems 8, "Scroll Up", "Scroll Left", "Scroll Right", "Scroll Down", "Click Left", "Click Middle", "Click Right", "Move Up", "Move Left", "Move Right", "Move Down"
                    End Select
                Case 6 'cancel
                    LCAR_AddListItems 8, "No", "Yes"
                Case 7 'save
                    SaveButton
            End Select
            LCARlists(8).Visible = LCARlists(8).ListCount > 0
            If LCARlists(8).Visible Then LCAR_SelectItem 8, LCAR_FindListItemByName(8, LCARlists(7).ListItems(Item).Side)

        Case 8 'button menu
            With LCARlists(7).ListItems(LCARlists(7).SelectedItem)
                .Side = LCARlists(8).ListItems(Item).Text
                .LightColor = LCARlists(8).ListItems(Item).LightColor
                .Color = LCARlists(8).ListItems(Item).Color
                .Tag = LCARlists(8).ListItems(Item).Tag
            End With
            Select Case LCARlists(7).SelectedItem
                Case 6 'cancel
                    Select Case Item
                        Case 0 'No
                        Case 1 'Yes
                            HideAllLists 0
                            Operation = Empty
                    End Select
            End Select
            
        Case 9 'ip address
            ShowKeyboard LCARtext, "editip", True, 0, 255
        Case 10 'port/blank/connect/cancel
            Select Case Item
                Case 0 'port
                    ShowKeyboard LCARtext, "editport", True, 0, 99999
                Case 1 'blank
                    LCAR_Select 10, 0
                Case 2 'connect
                    HideAllLists 0
                    IPAddress = GetIPAddress
                    Port = GetIPAddress(True)
                    ValidDisconnect = False
                    EmergencyDisable = True
                    Sock.CloseSck
                    Sock.Connect IPAddress, Port
                Case 3 'cancel
                    HideAllLists 0
            End Select
        
        Case 11 'task
            SendData "switchto " & Filename
        
        Case 12 'file
            If Size = -1 Then LCAR_ClearList 12
            SendData "folder dir '" & Filename & "' 1"
    End Select
    If RAND Then PlayRandomSound 101, 102, 103, 106, 107, 108, 109 Else PlayRandomSound 110
End Sub

Public Function GetIPAddress(Optional Port As Boolean) As String
    If Port Then
        GetIPAddress = LCARlists(10).ListItems(0).FileLCAR
    Else
        With LCARlists(9)
            GetIPAddress = .ListItems(0).FileLCAR & "." & .ListItems(1).FileLCAR & "." & .ListItems(2).FileLCAR & "." & .ListItems(3).FileLCAR
        End With
    End If
End Function

Private Sub Form_Load()
    Dim temp As Long
    'Set Dest = Me
    LCAR_AddDestination Me, "frmmain"
    'SetupVoiceRecognition
    Set buffer = Me.picbuffer
    Set Sock = New CSocketMaster
    Silent = True
    
    If Not IsFontInstalled("LCARS") Then
        MsgBox "Please install the LCAR font that was included with this program", vbCritical, "LCAR font missing"
        End
    End If
    
    SetupLCARcolors
    
    SetupEffects
    Mute = CBool(GetSetting("LCAR", "MAIN", "MUTE", "False"))
    SizeMode = CBool(GetSetting("LCAR", "MAIN", "SizeMode", "False"))
    
    LCAR_AddLCAR "frmmain", 2, 2, 120, 60, 100, 15, , , 2, , , 1, , , , False
    LCAR_AddLCAR "frmmain", 2, 64, 120, 60, 100, 15, , , 0, , , 2, , , 2, False
    
    LCAR_AddMenu "btnmenu", 1, 123, 47, 6, 101, 15, , 0, 0
    LCAR_SetTexts "btnmenu", 1, False, False, "Connect to server", "Options", "About", Empty, "_", "X"
    temp = LCAR_FindLCAR("btnmenu", 1, 1) 'options
    LCAR_ButtonList(temp).Width = 50
    temp = LCAR_FindLCAR("btnmenu", 1, 2) 'about
    LCAR_ButtonList(temp).X = LCAR_ButtonList(temp).X - 2
    LCAR_ButtonList(temp).Width = 54
    LCAR_ButtonList(temp).X = LCAR_ButtonList(temp).X - 49
    temp = LCAR_FindLCAR("btnmenu", 1, 3) 'empty
    LCAR_ButtonList(temp).X = LCAR_ButtonList(temp).X - 98 '_
    LCAR_ButtonList(temp).Width = -LCAR_ButtonList(temp).X - 66
    LCAR_ButtonList(temp).Enabled = False
    LCAR_ButtonList(temp).TextAlign = 6
    timebase = temp
    temp = LCAR_FindLCAR("btnmenu", 1, 4) '_
    LCAR_ButtonList(temp).X = -64
    LCAR_ButtonList(temp).Width = 30
    LCAR_ButtonList(temp).TextAlign = 2
    temp = LCAR_FindLCAR("btnmenu", 1, 5) 'X
    LCAR_ButtonList(temp).X = -32
    LCAR_ButtonList(temp).Width = 30
    LCAR_ButtonList(temp).TextAlign = 5
    
    '                                   items
    LCAR_AddMenu "btnpath", 2, 123, 64, 5, 10, 15, , 0, 0
    temp = LCAR_SetTexts("btnpath", 2, True, False, "New Profile", "Load Profile", "Delete Profile")
    LCAR_ButtonList(LCAR_FindLCAR("btnpath", 2, 3)).Width = -344
    temp = LCAR_FindLCAR("btnpath", 2, 4)
    LCAR_ButtonList(temp).Width = 62
    LCAR_ButtonList(temp).X = -64
    LCAR_ButtonList(temp).TextAlign = 2
    
    LCAR_AddMenu "btntasks", 3, 2, 126, 8, 100, 15, False, 0, 0
    temp = LCAR_SetTexts("btntasks", 3, False, False, "Keyboard", "Numboard", "Add Button", "Edit Button", "Delete Button", "Xperia Play", "Task List", "File List") ', "Open With", "Rename", "Copy", "Cut", "Paste", "Copy to", "Move to", "Delete", "Undo")
    
    temp = LCAR_NextY("btntasks", 3, temp)
    LCAR_AddLCAR "btnsearch", 2, temp, 100, -104 - temp, 0, 0, , , , "Stand Down", , 3, , , 2
    
    LCAR_AddMenu "btnsystem", 3, 2, -102, 6, 100, 15, False, 0, 0, , LCAR_ABC
    LCAR_SetTexts "btnsystem", 3, False, False, "Multi Select", "Select None", "Select All", "Invert Selected", "", "Rotate Screen"
    
    'LCAR_AddLCAR "btnlcar", 110, 90, 120, 20, 20, 0, LCAR_Orange, LCAR_DarkOrange, , "This is a button"
    'temp = LCAR_AddLCAR("btnlcar", 110, 112, 120, 20, 20, 0, LCAR_Orange, LCAR_DarkOrange, , "This is a blinkie")
    'LCAR_ButtonList(temp).State = -1
    
    'LISTS
    LCAR_AddList "lstfiles", 2, 3, 105, 82, -104, -84 '0
    LCAR_AddList "lstfolders", 1, 1, 105, 82, 200, -84, False '1
    
    'KEYBOARD CODE
    LCAR_AddList "KBTop", 10, 10, 2, -157, -4, 96, False, 0, 20, 20, 0, False '2
    'AddListItems 2, "`", "\", "'", "Shift", "Symb", LCAR_SMB, LCAR_SMB, LCAR_SMB, LCAR_CMD, LCAR_CMD
    AddListItems 2, "1", "Q", "A", "Z", LCAR_123, LCAR_ABC, LCAR_ABC, LCAR_ABC, "!", "~", "`"
    AddListItems 2, "2", "W", "S", "X", LCAR_123, LCAR_ABC, LCAR_ABC, LCAR_ABC, "@", "[", "{"
    AddListItems 2, "3", "E", "D", "C", LCAR_123, LCAR_ABC, LCAR_ABC, LCAR_ABC, "#", "]", "}"
    AddListItems 2, "4", "R", "F", "V", LCAR_123, LCAR_ABC, LCAR_ABC, LCAR_ABC, "$"
    AddListItems 2, "5", "T", "G", "B", LCAR_123, LCAR_ABC, LCAR_ABC, LCAR_ABC, "%", "|"
    AddListItems 2, "6", "Y", "H", "N", LCAR_123, LCAR_ABC, LCAR_ABC, LCAR_ABC, "^", "\"
    AddListItems 2, "7", "U", "J", "M", LCAR_123, LCAR_ABC, LCAR_ABC, LCAR_ABC, "&", "_"
    AddListItems 2, "8", "I", "K", ",", LCAR_123, LCAR_ABC, LCAR_ABC, LCAR_SMB, "*", "-"
    AddListItems 2, "9", "O", "L", ".", LCAR_123, LCAR_ABC, LCAR_ABC, LCAR_SMB, "(", "+"
    AddListItems 2, "0", "P", ";", "/", LCAR_123, LCAR_ABC, LCAR_SMB, LCAR_SMB, ")", "=", ":", "?"
    
    LCAR_AddLCAR "frmbottom", 2, -219, 125, 60, 100, 20, , , 0, "HOME", , 5, , , 2 '0
    LCAR_AddLCAR "frmbottom", 128, -219, 115, 20, 0, 0, , , , "LEFT", , 5, , , 5 '1
    LCAR_AddLCAR "frmbottom", Width / 2, -219, 100, 20, 0, 0, , , , "RIGHT", , 5, , , 5 '2
    LCAR_AddLCAR "frmbottom", -127, -219, 125, 60, 100, 20, , , 1, "END", , 5, , , 2 '3
    
    'LCAR_AddLCAR "frmbottom", -127, 81, 125, -302, 100, 20, , , 3, "SYMBOLS", , 5, , , 2   '4 +51 to Y, -51 to height
    LCAR_AddLCAR "frmbottom", -127, 149, 125, -370, 100, 20, , , 3, "SYMBOLS", , 5, , , 2    '4
    LCAR_AddLCAR "frmbottom", 2, 126, 100, -409, 0, 0, , , , "CAPS LOCK", , 5, , , 5  '5
    LCAR_AddLCAR "frmbottom", 2, -281, 125, 60, 100, 20, , , 2, "SHIFT", , 5, , , 2 '6
    
    LCAR_AddLCAR "frmbottom", 128, -241, 117, 20, 0, 0, , , , "DELETE", , 5, , , 5   '7
    LCAR_AddLCAR "frmbottom", Width / 2, -241, 100, 20, 0, 0, , , , "BACKSPACE", , 5, , , 5 '8
    
    LCAR_AddLCAR "frmbottom", 2, -62, 125, 60, 100, 20, , , 2, "CANCEL", , 5, , , 2 '9
    LCAR_AddLCAR "frmbottom", -127, -62, 125, 60, 100, 20, , , 3, "OK", , 5, , , 2 '10
    LCAR_AddLCAR "frmbottom", 128, -22, -257, 20, 0, 0, , , , "SPACE BAR", , 5, , , 5  '11
    
    GroupList(5).Visible = False
    
    'LISTS Continued...
    LCAR_AddList "lstprograms", 2, 3, 105, 82, -104, -84, False, , , , 50  '3
    
    LCAR_AddList "lstoptions", 1, 1, 105, 82, -305, -84, False, , , , 200    '4           options
    LCAR_AddList "lstoptionmenu", 1, 1, -199, 82, 200, -84, False, , , , 50    '5         options context menu
        LCAR_AddListItem 4, "Sound", , , , , , , IIf(Mute, "Off", "On")
        LCAR_AddListItem 4, "Display authentic file size", , , , , , , IIf(SizeMode, "On", "Off")
        LCAR_AddListItem 4, "Font size", , , , , , , GetSetting("LCAR", "MAIN", "FontSize", "0")
        LCAR_AddListItem 4, "UI mode", , , , , , , GetSetting("LCAR", "MAIN", "UI", "Classic") '  IIf(AntiAliasing, "Off", "On")
        
        
    LCAR_AddLCAR "btnsearch", 0, 64, 10, 15, 0, 0, , , , "Manual Entry", , 6, , , 5

    
    GroupList(6).Visible = False
    
    
    'red alert prompt
    LCAR_AddLCAR "btnaccessdenied", 106, 92, -108, 64, 0, 0, vbBlack, vbBlack, , "ACCESS DENIED", , 7, True, , 5, , 64
    LCAR_AddLCAR "btnaccessprompt", 106, 157, -108, 32, 0, 0, vbBlack, vbBlack, , "This will cause damage to the file system", , 7, True, , 5, , 16
    
    LCAR_AddLCAR "btnaprompt", 106, 193, 160, 21, 25, 0, , , , "Belay order", , 7, , , , , 16
    LCAR_AddLCAR "btnaprompt", 270, 193, 160, 21, 25, 0, , , , "Override safeties", , 7, , , , , 16
    
    GroupList(7).Visible = False
    
    
    
    'NUMBOARD CODE
    LCAR_AddList "KBNUM", 3, 3, 2, -157, -4, 96, False, , , , 0, True       '6           numboard
    AddNumItems 7, 4, 1, 0, 8, 5, 2, 0, 9, 6, 3, 0
    
    LCAR_AddLCAR "frmnumbottom", 2, -219, 125, 60, 100, 20, , , 0, "HOME", , 8, , , 2   '0
    LCAR_AddLCAR "frmnumbottom", 128, -219, 115, 20, 0, 0, , , , "LEFT", , 8, , , 5  '1
    LCAR_AddLCAR "frmnumbottom", Width / 2, -219, 100, 20, 0, 0, , , , "RIGHT", , 8, , , 5   '2
    LCAR_AddLCAR "frmnumbottom", -127, -219, 125, 60, 100, 20, , , 1, "END", , 8, , , 2  '3
    
    LCAR_AddLCAR "frmnumbottom", -127, -281, 125, 60, 100, 20, , , 3, "Minus 10", , 8, , , 2       '4
    LCAR_AddLCAR "frmnumbottom", 2, 126, 100, -409, 0, 0, , , , "Plus 1", , 8, , , 5  '5
    LCAR_AddLCAR "frmnumbottom", 2, -281, 125, 60, 100, 20, , , 2, "Minus 1", , 8, , , 2    '6
    
    LCAR_AddLCAR "frmnumbottom", 128, -241, 117, 20, 0, 0, , , , "DELETE", , 8, , , 5   '7
    LCAR_AddLCAR "frmnumbottom", Width / 2, -241, 100, 20, 0, 0, , , , "BACKSPACE", , 8, , , 5 '8
    
    LCAR_AddLCAR "frmnumbottom", 2, -62, 125, 60, 100, 20, , , 2, "CANCEL", , 8, , , 2 '9
    LCAR_AddLCAR "frmnumbottom", -127, -62, 125, 60, 100, 20, , , 3, "OK", , 8, , , 2 '10
    'LCAR_AddLCAR "frmnumbottom", -102, 81, 100, -364, 0, 0, , , , "Plus 10", , 8, , , 5      '11    +51 to Y, -51 to height
    LCAR_AddLCAR "frmnumbottom", -102, 149, 100, -432, 0, 0, , , , "Plus 10", , 8, , , 5      '11
    
    LCAR_AddLCAR "frmnumbottom", 128, -22, 115, 20, 0, 0, , , , "Minus 1000", , 8, , , 5  '12
    LCAR_AddLCAR "frmnumbottom", 256, -22, 115, 20, 0, 0, , , , "Minus 100", , 8, , , 5  '13
    LCAR_AddLCAR "frmnumbottom", 312, -22, 115, 20, 0, 0, , , , "Plus 100", , 8, , , 5  '14
    LCAR_AddLCAR "frmnumbottom", 400, -22, 115, 20, 0, 0, , , , "Plus 1000", , 8, , , 5  '15
    
    GroupList(8).Visible = False
    
    
    preview = LCAR_AddLCAR("btntext", 123, 2, -104, 44, 0, 0, vbBlack, vbBlack, -1, NoneSelected, , 0, True, , 1, False)
    TextID = LCAR_AddLCAR("btntxtbox", 112, 2, -104, 44, 0, 0, vbBlack, vbBlack, -1, NoneSelected, , 5, True, , 1, False, 30)
    LCAR_AddLCAR "btnnumbox", 112, 2, -104, 44, 0, 0, vbBlack, vbBlack, -1, NoneSelected, , 8, True, , 1, False, 30
    
    
    'add button form
    
    LCAR_AddList "lstbutton", 1, 1, 105, 82, -305, -84, False, , , , 200    '7                      new/edit button options
    LCAR_AddList "lstbuttonmenu", 1, 1, -199, 82, 200, -84, False, , , , 50    '8                   new/edit button context menu
       
    UMR_LoadHINI Hinimain
    LCAR_AddList "lstipaddress", 4, 4, 105, 82, -109, 40, False, , , , 0, True                  '9          selecting IP address
        LCAR_AddListItem 9, "•", LCAR_123, , , "0", , , , , "255", UMR_GetSetting("Defaults", "IP0", "255")
        LCAR_AddListItem 9, "•", LCAR_123, , , "1", , , , , "255", UMR_GetSetting("Defaults", "IP1", "255")
        LCAR_AddListItem 9, "•", LCAR_123, , , "2", , , , , "255", UMR_GetSetting("Defaults", "IP2", "255")
        LCAR_AddListItem 9, Empty, LCAR_123, , , "3", , , , , "255", UMR_GetSetting("Defaults", "IP3", "255")
    LCAR_AddList "lstipaddress2", 4, 4, 105, 106, -109, 40, False, , , , 0, True   '10     selecting port and connect/cancel
        LCAR_AddListItem 10, "Port", LCAR_ABC, , , , , , , , "21", UMR_GetSetting("Defaults", "Port", "21")
        LCAR_AddListItem 10, Empty, vbBlack
        LCAR_AddListItem 10, "Connect", LCAR_CMD
        LCAR_AddListItem 10, "Cancel", LCAR_CMD
    
    LCAR_AddList "lsttasks", 1, 1, 105, 82, -104, -84, False '11     used for task bar
    LCAR_AddList "lstthefiles", 2, 3, 105, 82, -104, -84 '12         used for file browsing
    
    
    LCAR_AddMenu "btnmouse", 9, -102, 81, 4, 100, 15, False, 0, 0
    LCAR_SetTexts "btnmouse", 9, False, False, "Left", "Middle", "Right", "Scroll"
    LCAR_Blink LCAR_FindLCAR("btnmouse", 9, 0), True
    Button = "left"
    GroupList(9).Visible = False
    
    
    swTolerance = UMR_GetSetting("Defaults", "Tolerance", "5")
    DebugMode = CBool(UMR_GetSetting(Empty, "DebugMode", "False"))
    Rotate = CBool(GetSetting("LCAR", "Main", "ROTATE", "False"))
    oldsize = Me.Font.Size
    
    If IsInIDE Then
        WindowState = vbNormal
        Me.Move 0, 0, 800 * 15, 480 * 15
    End If
    
    LCAR_FontIncrement Val(GetSetting("LCAR", "MAIN", "FontSize", "0"))
    SetupUImode GetSetting("LCAR", "MAIN", "UI", "Classic")
    
    Sock.PrintDebug = False
    Sock.AutoBind Val(UMR_GetSetting("Defaults", "LocalPort", "21")), True, True
    
    LCAR_AddListItem 4, "Telnet Port", , , , , , , CStr(Sock.LocalPort)
    LCAR_AddListItem 4, "Mouse Tolerance", , , , , , , UMR_GetSetting("Defaults", "Tolerance", "5")
    
    LCAR_ButtonList(LCAR_FindLCAR("btnpath", 2, 4)).Text = Sock.LocalIP
    Silent = True
    
    LCAR_ANIIncrementLocs , , True
End Sub

Public Sub ResetButtonOptions(Optional ButtonName As String)
        LCAR_ClearList 7
        
        If Len(ButtonName) > 0 Then
            LCAR_AddListItem 7, "Name:", , , , , , , ButtonName
        Else
            LCAR_AddListItem 7, "Name:", , , , , , , "New Button"
        End If
        
        LCAR_AddListItem 7, "Color:", , Val(UMR_ButtonProperty(ButtonName, "ColorValue", "-1")), , , , , UMR_ButtonProperty(ButtonName, "Color", "Light Blue")
        LCAR_AddListItem 7, "Action type:", , , , , , , UMR_ButtonProperty(ButtonName, "ActionType", "Keyboard")
        LCAR_AddListItem 7, "CTRL:", LCAR_SMB, , , , , , UMR_ButtonProperty(ButtonName, "CTRL", "No")
        LCAR_AddListItem 7, "ALT:", LCAR_SMB, , , , , , UMR_ButtonProperty(ButtonName, "ALT", "No")
        LCAR_AddListItem 7, "Button:", LCAR_ABC, , , , , , UMR_ButtonProperty(ButtonName, "Button", "None")

        LCAR_AddListItem 7, "Cancel", , , , , , , "No"
        LCAR_AddListItem 7, "Save"
End Sub
Public Sub SaveButton()
    Dim ButtonName As String, temp As Long, Fail As Boolean, Delete As Boolean
    ButtonName = UCase(LCARlists(7).ListItems(0).Side)
    
    If Len(OldButton) = 0 Then
        If UMHini.SectionExists("Profiles\" & CurrentProfile & "\" & ButtonName) Then Fail = True
    Else
        Delete = True
        If StrComp(OldButton, ButtonName, vbTextCompare) <> 0 Then
            If Not UMHini.RenameSection("Profiles\" & CurrentProfile & "\" & OldButton, ButtonName) Then
                Delete = False
                Fail = True
            End If
        End If
    End If
    If Fail Then
        Prompt "Duplicate detected, unable to use '" & ButtonName & "'", False, True
        Exit Sub
    End If
    If Delete Then
        temp = LCAR_FindListItemByName(0, OldButton)
        LCAR_DeleteListItem 0, temp
    End If
    
    LCAR_AddListItem 0, ButtonName, , LCARlists(7).ListItems(1).LightColor
    
    UMR_SaveButton CurrentProfile, ButtonName, LCARlists(7).ListItems(1).LightColor, LCARlists(7).ListItems(1).Side, LCARlists(7).ListItems(2).Side, LCARlists(7).ListItems(3).Side, LCARlists(7).ListItems(4).Side, LCARlists(7).ListItems(5).Side
    If Sock.IsTheServer Then
        If StrComp(CurrentWindow(Me.hwnd), CurrentProfile, vbTextCompare) = 0 Then
            SendData "addbutton """ & ButtonName & """ " & CStr(LCARlists(7).ListItems(1).LightColor)
        End If
    End If
    'UMR_ButtonProperty ButtonName, "ColorValue", CStr(LCARlists(7).ListItems(1).LightColor), , True
    'UMR_ButtonProperty ButtonName, "Color", LCARlists(7).ListItems(1).Side, , True
    'UMR_ButtonProperty ButtonName, "ActionType", LCARlists(7).ListItems(2).Side, , True
    'UMR_ButtonProperty ButtonName, "CTRL", LCARlists(7).ListItems(3).Side, , True
    'UMR_ButtonProperty ButtonName, "ALT", LCARlists(7).ListItems(4).Side, , True
    'UMR_ButtonProperty ButtonName, "Button", LCARlists(7).ListItems(5).Side, , True
    
    HideAllLists 0
End Sub



Public Sub SensorSweep_MouseDown(X As Long, Y As Long)
    'Debug.Print "Sensor Down " & X & ", " & Y & " " & swActive
    'GNDN
End Sub
Public Sub SensorSweep_MouseMove(X As Long, Y As Long)
    'Debug.Print "Sensor Move " & X & ", " & Y & " " & swActive
    If swActive And Sock.IsTheClient Then
        If Scrolling Then
            If Y <> 0 Then SendData "mouse scroll " & Y
            If X <> 0 Then SendData "mouse scrollx " & X
        Else
            SendData "mouse " & X & " " & Y
        End If
    End If
End Sub
Public Sub SensorSweep_MouseUp(X As Long, Y As Long)
    'Debug.Print "Sensor Up " & X & ", " & Y & " " & swActive
    If Not swActive And Sock.IsTheClient Then
        SendData "mouse " & Button
    End If
End Sub












Public Sub LCARMouseDown()
    'PlayRandomSound
   ' Debug.Print OldClickedAtX & ", " & clickedaty
End Sub
Public Sub LCARMouseUp()
    'PlayRandomSound
End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)
    Dim tempstr As String
    
    If KeyboardIsVisible Then
        Select Case KeyCode
            Case 16 'shift
                tempstr = "shift"
            Case Else
                tempstr = vKey2String(CLng(KeyCode))
                If (isShift And Caps) Or ((Not isShift) And (Not Caps)) Then tempstr = LCase(tempstr)
                If (Caps And Not isShift) Or (Not Caps And isShift) Then tempstr = UCase(tempstr)
        End Select
    
        ProcessKey tempstr
    End If
End Sub



Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim temp As Long, temp2 As Long ', tempstr As String, tempstr2 As String, temp3 As Long
    If Button <> vbLeftButton Then Exit Sub
    temp = LCAR_FindClicked(CLng(X), CLng(Y), True)
    LCARid = -1
    If temp > -1 Then
        If LCAR_ButtonList(temp).Enabled Then
            LCARid = temp
            LCARname = LCAR_ButtonList(temp).Name
            LCARindex = LCAR_FindIndex(temp)
            LCARGroup = LCAR_ButtonList(temp).Group
            LCARblinking = LCAR_ButtonList(temp).State = -1
        
            LCAR_ButtonList(temp).State = 1
            LCAR_ButtonList(temp).IsClean = False
            LCAR_DrawLCARs
            RaiseEvent LCARMouseDown(LCARname, LCARindex)
            LCARMouseDown
            
            OldClickedAtX = ClickedAtX
        End If
    Else
        temp = LCAR_FindList(X, Y)
        If temp > -1 Then
            ListId = temp
            LCARname = LCARlists(temp).Name
            isDown = True
            temp2 = LCAR_FindListItem(X, Y)
            LCARlists(temp).isDown = True
            If temp2 = -1 Then
                LCARitem = -1
            Else
                LCARitem = temp2
                'MsgBox LCARlists(temp).ListItems(temp2).Tag
                LCAR_SelectItem temp, temp2
                RefreshPreview
                LockedOn = True
                oldRow = LCAR_ClickedRow(ListId, CLng(X), CLng(Y), False)
                oldCol = LCAR_ClickedCol(ListId, CLng(X), CLng(Y), False)
            End If
        Else
            OldX = CLng(X)
            OldY = CLng(Y)
            If IsInSensorSweep(OldX, OldY) Then
                swDown = True
                swActive = False
                SensorSweep_MouseDown OldX, OldY
            End If
        End If
    End If
    
    'End If
End Sub



Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim temp As Long, temp2 As Long
    If isDown Then
        
        temp = LCAR_ClickedRow(ListId, CLng(X), CLng(Y), True)
        If temp <> oldRow Then
            If LockedOn Then
                LockedOn = False
                LCAR_SelectItem ListId, LCARlists(ListId).SelectedItem
                RefreshPreview
            End If
            LCAR_ScrollList ListId, oldRow - temp
            oldRow = temp
        End If
    ElseIf swDown Then
        temp = X - OldX
        temp2 = Y - OldY
        If Abs(temp) > swTolerance Or Abs(temp2) > swTolerance Then swActive = True
        If swActive Then
            OldX = X
            OldY = Y
            If temp <> 0 Or temp2 <> 0 Then
                If Rotate Then
                    SensorSweep_MouseMove -temp2, temp
                Else
                    SensorSweep_MouseMove temp, temp2
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim temp As Long
    If Button <> vbLeftButton Then Exit Sub
    
    If isDown Then
        isDown = False
        LCARlists(ListId).isDown = False
        If LockedOn And LCARitem > -1 Then LCARListItemClicked
    Else

    
            temp = LCAR_FindClicked(CLng(X), CLng(Y), True)
            If LCARid > -1 Then
                RaiseEvent LCARMouseUp(LCARname, LCARindex)
                LCARMouseUp
                LCAR_ButtonList(LCARid).State = IIf(LCARblinking, -1, 0)
                LCAR_ButtonList(LCARid).IsClean = False
    
                If temp = LCARid Then
                    RaiseEvent LCARClicked(LCARname, LCARindex)
                    LCARClicked
                End If
            ElseIf swDown Then
                SensorSweep_MouseUp CLng(X), CLng(Y)
                swDown = False
            End If
    
    End If
    
    LCAR_DrawLCARs
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    UMR_SaveHINI
    SaveSetting "LCAR", "Main", "Rotate", Rotate
    SendData "disconnect"
    Sock.CloseSck
    End
End Sub

Private Sub Form_Resize()
    If WindowState = vbMinimized Then
        wasrotated = Rotate
        wasminimized = True
        IsInFocus = False
    ElseIf wasminimized Then
        IsInFocus = True
        wasminimized = False
        If wasrotated Then SwitchToUnRotated 'RotateScreen: RotateScreen: RotateScreen
    End If
    IsClean = False
    ResizeLCARs
    LCAR_DrawLCARs True
End Sub


Private Sub Sock_Connect()
    EmergencyDisable = False
    SOCK_ConnectionRequest 0
End Sub


Public Sub TimerEffects_Timer()
    If IsInFocus Then
        DrawEffects
    Else
        HandleGameMode
    End If
End Sub














Private Sub Sock_Error(ByVal Number As Integer, Description As String, ByVal sCode As Long, ByVal Source As String, ByVal HelpFile As String, ByVal HelpContext As Long, CancelDisplay As Boolean)
    On Error Resume Next
    Prompt GetErrorDescription(Number), False
    Sock.Listen
    LCARlists(4).ListItems(4).Side = Sock.LocalPort
End Sub
Private Sub SOCK_CloseSck()
    GameMode = Empty
    If Client Then
        Prompt "Override to attempt reconnection?", True, True, "DISCONNECTED", "reconnect"
    Else
        Sock.Listen
        LCARlists(4).ListItems(4).Side = Sock.LocalPort
    End If
End Sub

Private Sub TimerBlink_Timer()
    Dim tempstr As String
    tempstr = CurrentWindow(Me.hwnd)
    If StrComp(tempstr, OldWindow, vbTextCompare) <> 0 And Len(tempstr) > 0 Then
        OldWindow = tempstr
        If Sock.IsTheServer Then SOCK_ConnectionRequest 0
    End If
    
    If IsInFocus Then
        With LCAR_ButtonList(timebase)
            .Text = IIf(Rotate, "S: ", "STARDATE: ") & StarDate(Now, 5)
            .IsClean = False
        End With
        LCAR_BlinkLCARs
        If Not IsClean Then Form_Resize ' LCAR_DrawLCARs True
        
        DrawCondition 150, 100, 400, vbRed
    End If
End Sub

Private Sub SOCK_ConnectionRequest(ByVal requestID As Long)
    Dim tempstr As String
    OldWindow = CurrentWindow(Me.hwnd)
    tempstr = OldWindow
    If Not UMR_ProfileExists(tempstr) Then tempstr = "Default"
    SendData "Profile """ & tempstr & """ " & UMR_StarDate(OldWindow)
    SendData "connected"
    UMR_LoadProfile tempstr, Sock
End Sub

Sub SendData(Data As String)
    Sock.SendData Data & Chr(10)
End Sub

Private Sub SOCK_DataRecieved(ByVal Data As String)
    Dim temp As Long, tempstr() As String, Count As Long, Word As String, char As String
    If Len(Data) = 0 Or Data = Chr(10) Then Exit Sub
    'Call MsgBox(Data, , Len(Data) & Asc(Right(Data, 1)))
    If InStr(Data, Chr(10)) Then
        tempstr = Split(Data, Chr(10))
        For temp = 0 To UBound(tempstr)
            'MsgBox tempstr(temp)
            SOCK_DataRecieved tempstr(temp)
        Next
        Exit Sub
    End If
    
    Count = Sock.SplitCommand(Data, tempstr, , True)
    
    If DebugMode Then RefreshPreview Data
    'Debug.Print "Data: <" & Data & ">"
    'For temp = 0 To count - 1
        'Debug.Print "Word #" & temp & "/" & count & ": '" & tempstr(temp) & "'"
    'Next
    
    
If Count > 0 Then
    Select Case LCase(tempstr(0))
        'GET is reserved by the Web server component
        Case "g" 'gamepad data
            Word = tempstr(1)
            If Len(Word) > 15 Then Word = Left(Word, 15)
            GameMode = Word
            Debug.Print "RECEIVED: " & Word
            HandleGameMode
        
        Case "raw"
            If Not rawmode Then SetupVoiceRecognition
            rawmode = True
            
            
        Case "disconnect"
            ValidDisconnect = True
            If Client Then
                Prompt "The client has shut down", False, False, "DISCONNECTED"
            Else
                Prompt "The server has shut down", False, True, "DISCONNECTED"
            End If
            
        Case "profile"
            If Sock.IsTheServer Then
                'send profile to client
                UMR_LoadProfile tempstr(1), Sock
            Else
                'load profile, if not found or is older than the stardate received, then request the profile
                CurrentProfile = tempstr(1)
                LCAR_SetText "btnpath", "Current profile: " & CurrentProfile, 3, 2
                'If UMR_StarDate(CurrentProfile) < CDbl(tempstr(2)) Then 'if older or not found (not found will always be older)
                    UMR_DeleteProfile CurrentProfile, -1
                    UMR_NewProfile CurrentProfile, -1
                    'SendData "Profile """ & CurrentProfile
                'Else 'current data is up to date, load it
                '    UMR_LoadProfile CurrentProfile
                'End If
            End If

        Case "addbutton" 'addbutton name color
            If Sock.IsTheServer Then
                Debug.Print "add button from client to profile"
            Else 'add button from server to list/screen
                UMR_SaveButton CurrentProfile, tempstr(1), Val(tempstr(2))
            End If
        Case "requestbutton"
            If Sock.IsTheServer Then 'client is requesting more button data
                UMR_ButtonData tempstr(1), tempstr(2), Sock
            Else 'server is sending more button data for editing
            
            End If
        
        Case "execute" 'execute profile button
            'client has told the server which profile\button to execute
            'the server cannot tell the client to do this
            UMR_Execute GetParam(tempstr, 2, CurrentProfile), tempstr(1)
            
        Case "task" 'task [hwnd]
            If Sock.IsTheServer Then
                UMR_SendTasks Sock
            Else
                LCAR_AddListItem 11, tempstr(2), , , , tempstr(1)
            End If
        Case "switchto"
            SetForegroundWindow Val(tempstr(1))
            'OldWindow = CurrentWindow(Me.hwnd)
            TimerBlink_Timer
            
            
        Case "folder" 'folder dir|path path
            
            If Sock.IsTheServer Then
                'If StrComp(tempstr(1), "dir", vbTextCompare) = 0 Then
                    If UMR_SendFiles(Sock, GetParam(tempstr, 2), drvmain, Dirmain, Filmain) Then
                'Else
                        ShellFile Me.hwnd, tempstr(2)
                    End If
                'End If
            Else
                If Right(Data, 1) <> "'" Then
                    SOCK_DataRecieved Data & "'"
                Else
                    If StrComp(tempstr(1), "dir", vbTextCompare) = 0 Then ', "Folder"
                        If GetParam(tempstr, 3, "1") = "0" Then
                            LCAR_AddListItem 12, "..", -1, -1, -1, tempstr(2), -1, False
                        Else
                            LCAR_AddListItem 12, FileTitle(tempstr(2)), -1, -1, -1, tempstr(2), -1, False
                        End If
                    Else 'File
                        LCAR_AddListItem 12, FileTitle(tempstr(2), True), -1, -1, CLng(tempstr(1)), tempstr(2), -1, False, GetExtention(tempstr(2)) 'File
                    End If
                End If
            End If
        
        Case "mouse"
            Select Case LCase(tempstr(1))
                Case "left":    Mouse_Click vbLeftButton
                Case "right":   Mouse_Click vbRightButton
                Case "middle":  Mouse_Click vbMiddleButton
                Case "scroll":  Mouse_Scroll -Val(tempstr(2))
                Case "scrollx": Mouse_Scroll -Val(tempstr(2)), True
                Case Else:      Mouse_MoveTo Val(tempstr(1)), Val(tempstr(2))
            End Select
            'Debug.Print tempstr(1)
            
        Case "connected"
            RefreshPreview "You are now connected!"
            
        Case "sendtext"
            SendKeys Right(Data, Len(Data) - 9) 'tempstr(1)
        
        Case "setdir"
            CurrentDir = Right(Data, Len(Data) - 7)
        Case "exe"
            UMR_Execute Right(Data, Len(Data) - 4)
    End Select
End If

End Sub
Public Function GetParam(tempstr, Optional Parameter As Long, Optional Default As String) As String
    If Parameter >= 0 And Parameter <= UBound(tempstr) Then
        GetParam = tempstr(Parameter)
    Else
        GetParam = Default
    End If
End Function



Private Sub Sock_HTTPRequest(Filename As String, Request As String)
    Dim QueryList() As String, QueryCount As Long, tempstr As String
    If Len(Request) > 0 Then SOCK_DataRecieved Request
    tempstr = UMR_ProcessTemplate(Filename, drvmain, Dirmain, Filmain)
    Sock.HTTPReplyWithString tempstr
    Clipboard.Clear
    Clipboard.SetText tempstr
    'Sock.HTTPReplyWithString (Filename & " was requested with '" & Request & "'")
End Sub
