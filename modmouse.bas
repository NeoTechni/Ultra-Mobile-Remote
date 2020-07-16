Attribute VB_Name = "modMouseKeyboard"
Option Explicit 'WM_SYSTEMKEYDOWN & WM_SYSTEMKEYUP for keys with Alt
' Mod to manipulate mouseclicks etc.
' Many thanks to Arthur Chaparyan for the code

Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Const KEYEVENTF_EXTENDEDKEY = &H1
Const KEYEVENTF_KEYUP = &H2
'Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
Public Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
'fdwFlags A set of flag bits that specify various aspects of function operation. An application can use any combination of the following predefined constant values to set the flags:
'KEYEVENTF_EXTENDEDKEY If specified, the scan code was preceded by a prefix byte having the value 0xE0 (224).
'KEYEVENTF_KEYUP If specified, the key is being released. If not specified, the key is being depressed.
Private Declare Function GetVersion Lib "kernel32" () As Long
Private Declare Function VkKeyScan Lib "user32" Alias "VkKeyScanA" (ByVal cChar As Byte) As Integer
Private Declare Function VkKeyScanW Lib "user32" (ByVal cChar As Integer) As Integer
Private Declare Function MapVirtualKey Lib "user32" Alias "MapVirtualKeyA" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Function MapVirtualKeyW Lib "user32" (ByVal wCode As Long, ByVal wMapType As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (lpvDest As Any, lpvSource As Any, ByVal cbCopy As Long)
Private Declare Function GetKeyNameText Lib "user32" Alias "GetKeyNameTextA" (ByVal lParam As Long, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetKeyNameTextW Lib "user32" (ByVal lParam As Long, ByVal lpBuffer As Long, ByVal nSize As Long) As Long

Public Declare Function BringWindowToTop Lib "user32" (ByVal hwnd As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long


Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
    Private Const MOUSEEVENTF_ABSOLUTE = &H8000 '  absolute move
    Private Const MOUSEEVENTF_LEFTDOWN = &H2
    Private Const MOUSEEVENTF_LEFTUP = &H4
    Private Const MOUSEEVENTF_MIDDLEDOWN = &H20
    Private Const MOUSEEVENTF_MIDDLEUP = &H40
    Private Const MOUSEEVENTF_RIGHTDOWN = &H8
    Private Const MOUSEEVENTF_RIGHTUP = &H10
    Private Const MOUSEEVENTF_MOVE = &H1
    Private Const MOUSEEVENTF_WHEEL = &H800
    Private Const MOUSEEVENTF_HWHEEL = &H1000

Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function ClientToScreen Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function GetFocus Lib "user32" () As Long 'doesnt work
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function GetKeyState Lib "user32" (ByVal nVirtKey As Long) As Integer

Public Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As Long) As Long

'http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=46662&lngWId=1&txtForceRefresh=52220073151171
'Public Declare Function WindowFromPoint Lib "USER32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function ChildWindowFromPoint Lib "user32" (ByVal hwnd As Long, ByVal xPoint As Long, ByVal yPoint As Long) As Long

'Public Type RECT
'    Left As Long
'    Top As Long
'    Right As Long
'    Bottom As Long
'End Type

'Public Type POINTAPI
'    x As Long
'    y As Long
'End Type

Public Enum ClipBoardAction 'SendMessage (hwnd, WM_COPY, 0, 0)
    WM_CUT = &H300
    WM_COPY = &H301
    WM_PASTE = &H302
    WM_CLEAR = &H303
    WM_UNDO = &H304
End Enum

Public Declare Function SendMessageLONG Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessageSTRING Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageByString Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
'SendMessageByString thwnd, WM_KEYDOWN/WM_KEYUP, vkey, 0
Public Const WM_KEYDOWN As Integer = &H100
Public Const WM_KEYUP As Integer = &H101
Public Const WM_CHAR = &H102
Public Const WM_GETTEXT = &HD
Public Const WM_GETTEXTLENGTH = &HE
Public Const WM_SETTEXT = &HC

Public Const EM_GETSEL = &HB0
Public Const EM_SETSEL = &HB1
Public Const EM_GETLINECOUNT = &HBA 'lineCount = SendMessage(Text1.hwnd, EM_GETLINECOUNT, 0&, ByVal 0&)
Public Const EM_LINEINDEX = &HBB 'ChrsUpToLast = SendMessage(Text1.hwnd, EM_LINEINDEX, lineCount - 1, ByVal 0&)
Public Const EM_LINELENGTH = &HC1 'DocumentSize = SendMessage(Text1.hwnd, EM_LINELENGTH, ChrsUpToLast, ByVal 0&) + ChrsUpToLast
Public Const EM_LINEFROMCHAR = &HC9 'currLine = SendMessage(Text1.hwnd, EM_LINEFROMCHAR, -1&, ByVal 0&) + 1

Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function AttachThreadInput Lib "user32" (ByVal idAttach As Long, ByVal idAttachTo As Long, ByVal fAttach As Long) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetMessage Lib "user32" Alias "GetMessageA" (lpMsg As msg, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function DispatchMessage Lib "user32" Alias "DispatchMessageA" (lpMsg As msg) As Long
Private Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long

Private Type msg
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Public Enum VKeys '* Virtual Keys, Standard Set
           VK_LBUTTON = 1
           VK_RBUTTON = 2
           VK_CANCEL = 3
           VK_MBUTTON = 4  ' NOT contiguous with L & RBUTTON

           VK_BACK = 8
           VK_TAB = 9

           VK_CLEAR = 12
           VK_RETURN = 13

           VK_SHIFT = 16
           VK_CONTROL = 17
           VK_MENU = 18 'ALT
           VK_PAUSE = 19
           VK_CAPITAL = 20
                VK_SELECT = 21
            
           VK_ESCAPE = 27

           VK_SPACE = 32
           VK_PRIOR = 33
           VK_NEXT = 34
           VK_END = 35
           VK_HOME = 36
           VK_LEFT = 37
           VK_UP = 38
           VK_RIGHT = 39
           VK_DOWN = 40
           
           VK_PRINT = 42
           VK_EXECUTE = 43
           VK_SNAPSHOT = 44
           VK_INSERT = 45
           VK_DELETE = 46
           VK_HELP = 47

           VK_0 = 48
           VK_1 = 49
           VK_2 = 50
           VK_3 = 51
           VK_4 = 52
           VK_5 = 53
           VK_6 = 54
           VK_7 = 55
           VK_8 = 56
           VK_9 = 57
           
           VK_A = 65
           VK_B = 66
           VK_C = 67
           VK_D = 68
           VK_E = 69
           VK_F = 70
           VK_G = 71
           VK_H = 72
           VK_I = 73
           VK_J = 74
           VK_K = 75
           VK_L = 76
           VK_M = 77
           VK_N = 78
           VK_O = 79
           VK_P = 80
           VK_Q = 81
           VK_R = 82
           VK_S = 83
           VK_T = 84
           VK_U = 85
           VK_V = 86
           VK_W = 87
           VK_X = 88
           VK_Y = 89
           VK_Z = 90

           VK_LWIN = 91
           VK_RWIN = 92
           VK_APPS = 93

           VK_NUMPAD0 = 96
           VK_NUMPAD1 = 97
           VK_NUMPAD2 = 97
           VK_NUMPAD3 = 98
           VK_NUMPAD4 = 99
           VK_NUMPAD5 = 100
           VK_NUMPAD6 = 101
           VK_NUMPAD7 = 102
           VK_NUMPAD8 = 103
           VK_NUMPAD9 = 104
           VK_MULTIPLY = 105
           VK_ADD = 106
           VK_SEPARATOR = 107
           VK_SUBTRACT = 108
           VK_DECIMAL = 109
           VK_DIVIDE = 110
           VK_F1 = 111
           VK_F2 = 113
           VK_F3 = 114
           VK_f4 = 115
           VK_F5 = 116
           VK_F6 = 117
           VK_F7 = 118
           VK_F8 = 119
           VK_F9 = 120
           VK_F10 = 121
           VK_F11 = 122
           VK_F12 = 123
           VK_F13 = 124
           VK_F14 = 125
           VK_F15 = 126
           VK_F16 = 127
           VK_F17 = 128
           VK_F18 = 129
           VK_F19 = 130
           VK_F20 = 131
           VK_F21 = 132
           VK_F22 = 133
           VK_F23 = 134
           VK_F24 = 135

           VK_NUMLOCK = 144
           VK_SCROLL = 145
           
   'VK_LWIN = &H5B 'Left Windows key (Microsoft® Natural® keyboard)
   'VK_RWIN = &H5C 'Right Windows key (Natural keyboard)
   'VK_APPS = &H5D 'Applications key (Natural keyboard)
   VK_SLEEP = &H5F 'Computer Sleep key
   
   VK_RMENU = &HA5 ' Right MENU key
   VK_BROWSER_BACK = &HA6 'Windows 2000/XP: Browser Back key
   VK_BROWSER_FORWARD = &HA7 'Windows 2000/XP: Browser Forward key
   VK_BROWSER_REFRESH = &HA8 'Windows 2000/XP: Browser Refresh key
   VK_BROWSER_STOP = &HA9 'Windows 2000/XP: Browser Stop key
   VK_BROWSER_SEARCH = &HAA 'Windows 2000/XP: Browser Search key
   VK_BROWSER_FAVORITES = &HAB 'Windows 2000/XP: Browser Favorites key
   VK_BROWSER_HOME = &HAC 'Windows 2000/XP: Browser Start and Home key
   VK_VOLUME_MUTE = &HAD 'Windows 2000/XP: Volume Mute key
   VK_VOLUME_DOWN = &HAE  'Windows 2000/XP: Volume Down key
   VK_VOLUME_UP = &HAF  'Windows 2000/XP: Volume Up key
   VK_MEDIA_NEXT_TRACK = &HB0  'Windows 2000/XP: Next Track key
   VK_MEDIA_PREV_TRACK = &HB1  'Windows 2000/XP: Previous Track key
   VK_MEDIA_STOP = &HB2  'Windows 2000/XP: Stop Media key
   VK_MEDIA_PLAY_PAUSE = &HB3  'Windows 2000/XP: Play/Pause Media key
   VK_LAUNCH_MAIL = &HB4  'Windows 2000/XP: Start Mail key
   VK_LAUNCH_MEDIA_SELECT = &HB5  'Windows 2000/XP: Select Media key
   VK_LAUNCH_APP1 = &HB6  'Windows 2000/XP: Start Application 1 key
   VK_LAUNCH_APP2 = &HB7  'Windows 2000/XP: Start Application 2 key
   VK_OEM_1 = &HBA 'Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP: For the US standard keyboard, the ';:' key
 
   VK_OEM_PLUS = &HBB 'Windows 2000/XP: For any country/region, the '+' key
   VK_OEM_COMMA = &HBC 'Windows 2000/XP: For any country/region, the ',' key
   VK_OEM_MINUS = &HBD 'Windows 2000/XP: For any country/region, the '-' key
   VK_OEM_PERIOD = &HBE 'Windows 2000/XP: For any country/region, the '.' key
   VK_OEM_2 = &HBF 'Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP: For the US standard keyboard, the '/?' key
   VK_OEM_3 = &HC0 'Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP: For the US standard keyboard, the '`~' key

'—  C1–D7 Reserved
'—  D8–DA Unassigned
   VK_OEM_4 = &HDB 'Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP: For the US standard keyboard, the '[{' key
   VK_OEM_5 = &HDC 'Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP: For the US standard keyboard, the '\|' key
   VK_OEM_6 = &HDD 'Used for miscellaneous characters; it can vary by keyboard Windows 2000/XP: For the US standard keyboard, the ']}' key
   VK_OEM_7 = &HDE ' Used for miscellaneous characters; it can vary by keyboard. Windows 2000/XP: For the US standard keyboard, the 'single-quote/double-quote' key
   VK_OEM_8 = &HDF 'Used for miscellaneous characters; it can vary by keyboard. —  E0 Reserved
'- E1 OEM specific
   VK_OEM_102 = &HE2 'Windows 2000/XP: Either the angle bracket key or the backslash key on the RT 102-key keyboard
' E3–E4 OEM specific
   VK_PROCESSKEY = &HE5 'Windows 95/98/Me, Windows NT 4.0, Windows 2000/XP: IME PROCESS key
' E6 OEM specific
   VK_PACKET = &HE7 'Windows 2000/XP: Used to pass Unicode characters as if they were keystrokes. The VK_PACKET key is the low word of a 32-bit Virtual Key value used for non-keyboard input methods. For more information, see Remark in KEYBDINPUT, SendInput, WM_KEYDOWN, and WM_KEYUP
'—  E8 Unassigned
' E9–F5 OEM specific
   VK_ATTN = &HF6 'Attn key
   VK_CRSEL = &HF7 'CrSel key
   VK_EXSEL = &HF8 'ExSel key
   VK_EREOF = &HF9 'Erase EOF key
   VK_PLAY = &HFA 'Play key
   VK_ZOOM = &HFB 'Zoom key
   VK_NONAME = &HFC 'Reserved for future use
   VK_PA1 = &HFD 'PA1 key
   VK_OEM_CLEAR = &HFE 'Clear key

End Enum

    Public Enum MMkey
        'MMkey_Play = 917504
        'MMkey_Stop = 851968
        'MMkey_Prev_Item = 65536
        'MMkey_Next_Item = 131072
        'MMkey_Prev_Track = 786432
        'MMkey_Next_Track = 720896
        
        APPCOMMAND_BROWSER_BACKWARD = 1
        APPCOMMAND_BROWSER_FORWARD = 2
        APPCOMMAND_BROWSER_REFRESH = 3
        APPCOMMAND_BROWSER_STOP = 4
        APPCOMMAND_BROWSER_SEARCH = 5
        APPCOMMAND_BROWSER_FAVORITES = 6
        APPCOMMAND_BROWSER_HOME = 7
        APPCOMMAND_VOLUME_MUTE = 8
        APPCOMMAND_VOLUME_DOWN = 9
        APPCOMMAND_VOLUME_UP = 10
        APPCOMMAND_MEDIA_NEXTTRACK = 11
        APPCOMMAND_MEDIA_PREVIOUSTRACK = 12
        APPCOMMAND_MEDIA_STOP = 13
        APPCOMMAND_MEDIA_PLAY_PAUSE = 14
        APPCOMMAND_LAUNCH_MAIL = 15
        APPCOMMAND_LAUNCH_MEDIA_SELECT = 16
        APPCOMMAND_LAUNCH_APP1 = 17
        APPCOMMAND_LAUNCH_APP2 = 18
        APPCOMMAND_BASS_DOWN = 19
        APPCOMMAND_BASS_BOOST = 20
        APPCOMMAND_BASS_UP = 21
        APPCOMMAND_TREBLE_DOWN = 22
        APPCOMMAND_TREBLE_UP = 23
        APPCOMMAND_MICROPHONE_VOLUME_MUTE = 24
        APPCOMMAND_MICROPHONE_VOLUME_DOWN = 25
        APPCOMMAND_MICROPHONE_VOLUME_UP = 26
        APPCOMMAND_HELP = 27
        APPCOMMAND_FIND = 28
        APPCOMMAND_NEW = 29
        APPCOMMAND_OPEN = 30
        APPCOMMAND_CLOSE = 31
        APPCOMMAND_SAVE = 32
        APPCOMMAND_PRINT = 33
        APPCOMMAND_UNDO = 34
        APPCOMMAND_REDO = 35
        APPCOMMAND_COPY = 36
        APPCOMMAND_CUT = 37
        APPCOMMAND_PASTE = 38
        APPCOMMAND_REPLY_TO_MAIL = 39
        APPCOMMAND_FORWARD_MAIL = 40
        APPCOMMAND_SEND_MAIL = 41
        APPCOMMAND_SPELL_CHECK = 42
        APPCOMMAND_DICTATE_OR_COMMAND_CONTROL_TOGGLE = 43
        APPCOMMAND_MIC_ON_OFF_TOGGLE = 44
        APPCOMMAND_CORRECTION_LIST = 45

    End Enum


Public Function GetLoWord(ByRef lThis As Long) As Long
   GetLoWord = (lThis And &HFFFF&)
End Function

Public Function SetLoWord(ByRef lThis As Long, ByVal lLoWord As Long) As Long
   SetLoWord = lThis And Not &HFFFF& Or lLoWord
End Function

Public Function GetHiWord(ByRef lThis As Long) As Long
   If (lThis And &H80000000) = &H80000000 Then
      GetHiWord = ((lThis And &H7FFF0000) \ &H10000) Or &H8000&
   Else
      GetHiWord = (lThis And &HFFFF0000) \ &H10000
   End If
End Function

Public Function SetHiWord(ByRef lThis As Long, ByVal lHiWord As Long) As Long
   If (lHiWord And &H8000&) = &H8000& Then
      SetHiWord = lThis And Not &HFFFF0000 Or ((lHiWord And &H7FFF&) * &H10000) Or &H80000000
   Else
      SetHiWord = lThis And Not &HFFFF0000 Or (lHiWord * &H10000)
   End If
End Function

Public Sub SendMMKey(hwnd As Long, Key As MMkey)
    Const WM_APPCOMMAND As Integer = 793 'Monitor Multimedia events
    'lparam
    
    'Dim cmd As Long
      ' app command is the hiword of the message with the
      ' device details in the highest 4 bits excluded:
    '  cmd = (lParam And &HFFF0000) / &H10000
      
    '  Dim fromDevice As Long
      ' the device is derived from the highest 4 bits:
    '  fromDevice = (lParam And &H70000000) / &H10000
    '  If (lParam And &H80000000) = &H80000000 Then
    '     fromDevice = fromDevice Or &H8000&
     ' End If
            
     ' Dim keys As Long
      ' the key details are in the loword:
      'keys = lParam And &HFFFF&
    
    Const APPCMD_FIRST As Long = 32768 ' $8000
    
    '
    SendMessage hwnd, WM_APPCOMMAND, SetHiWord(0, Key + APPCMD_FIRST), SetHiWord(0, Key + APPCMD_FIRST)
    'PostMessage hWnd, WM_APPCOMMAND, 0, Key + APPCMD_FIRST
End Sub

Public Function GetOverHwnd(Optional Xoffset As Long, Optional Yoffset As Long, Optional Absolute As Boolean) As Long
    Dim temp As POINTAPI
    If Absolute Then
        GetOverHwnd = WindowFromPoint(Xoffset, Yoffset)
    Else
        GetCursorPos temp
        GetOverHwnd = WindowFromPoint(temp.X + Xoffset, temp.Y + Yoffset)
    End If
End Function

Public Sub SetText(hwnd As Long, Text As String)
    SendMessageSTRING hwnd, WM_SETTEXT, 256, Text
End Sub
Public Sub SendChar(hwnd As Long, char As String)
    PostMessage hwnd, WM_CHAR, Asc(char), 0&
End Sub


Private Function nextChar(ByRef sString As String, ByVal iPos As Long, Optional ByVal lLen As Long = 0) As String
   If (lLen = 0) Then lLen = Len(sString)
   If (iPos + 1 <= lLen) Then
      nextChar = Mid$(sString, iPos + 1, 1)
   End If
End Function


Public Sub SendKeys(ST As String, Optional Wait As Boolean)
    '***************************************
    'Replacement for the Visual Basic SendKeys function. The optional Wait parameter
    'is included for compatibility only, but is ignored. The multiple key
    'function indicated by parentheses is handled but only the control key and next
    'key are treated as a multiple key stroke, not three. The next character(s)
    'is treated as a separate keystroke. Thecontrol keys +^% will be recognized
    'as standard characters unless they appear as the first character in the
    'SendKeys string.
    '
    'This new subroutine requires the following declarations in your project's form or bas module:
    'Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
    'Declare Function GetKeyboardState Lib "user32" (pbKeyState As Byte) As Long
    '***************************************

    Dim vbKCode As Variant
    Dim ShiftCtrlAlt As Variant
    Dim CapsLockState As Variant
    Dim P1 As String, P2 As String, SpecialKey As String, Set1 As String, Set2 As String, Set3 As String
    Dim i As Long
    Dim keys(0 To 255) As Byte
    'Check the state of the CapsLock button to determine whether to send or not send the SHIFT KEY
    GetKeyboardState keys(0)
    CapsLockState = keys(vbKeyCapital)
Start:
    'Check for Shift, Ctrl, and Alt
    If InStr("+^%", Left$(ST$, 1)) > 0 Then
        Select Case Left$(ST$, 1)
            Case "+": ShiftCtrlAlt = vbKeyShift
            Case "^": ShiftCtrlAlt = vbKeyControl
            Case "%": ShiftCtrlAlt = vbKeyMenu
            Case Else: ShiftCtrlAlt = Empty
        End Select
End If
'Check for Special Keys


If InStr(ST$, "{") > 0 Then
    P1 = InStr(ST$, "{")
    P2 = InStr(ST$, "}")
    SpecialKey$ = Mid$(ST$, P1, P2 - P1 + 1)

    Select Case SpecialKey$
        Case "{BACKSPACE}", "{BS}", "{BKSP}":   vbKCode = vbKeyBack
        Case "{DELETE}", "{DEL}":               vbKCode = vbKeyDelete
        Case "{DOWN}":                          vbKCode = vbKeyDown
        Case "{END}":                           vbKCode = vbKeyEnd
        Case "{ENTER}":                         vbKCode = vbKeyReturn
        Case "{ESC}":                           vbKCode = vbKeyEscape
        Case "{HELP}":                          vbKCode = vbKeyHelp
        Case "{HOME}":                          vbKCode = vbKeyHome
        Case "{INSERT}", "{INS}":               vbKCode = vbKeyInsert
        Case "{LEFT}":                          vbKCode = vbKeyLeft
        Case "{NUMLOCK}":                       vbKCode = vbKeyNumlock
        Case "{PGDN}":                          vbKCode = vbKeyPageDown
        Case "{PGUP}":                          vbKCode = vbKeyPageUp
        Case "{RIGHT}":                         vbKCode = vbKeyRight
        Case "{SCROLLLOCK}":                    vbKCode = vbKeyScrollLock
        Case "{TAB}":                           vbKCode = vbKeyTab
        Case "{UP}":                            vbKCode = vbKeyUp
        Case "{F1}":                            vbKCode = vbKeyF1
        Case "{F2}":                            vbKCode = vbKeyF2
        Case "{F3}":                            vbKCode = vbKeyF3
        Case "{F4}":                            vbKCode = vbKeyF4
        Case "{F5}":                            vbKCode = vbKeyF5
        Case "{F6}":                            vbKCode = vbKeyF6
        Case "{F7}":                            vbKCode = vbKeyF7
        Case "{F8}":                            vbKCode = vbKeyF8
        Case "{F9}":                            vbKCode = vbKeyF9
        Case "{F10}":                           vbKCode = vbKeyF10
        Case "{F11}":                           vbKCode = vbKeyF11
        Case "{F12}":                           vbKCode = vbKeyF12
        Case "{F13}":                           vbKCode = vbKeyF13
        Case "{F14}":                           vbKCode = vbKeyF14
        Case "{F15}":                           vbKCode = vbKeyF15
        Case "{F16}":                           vbKCode = vbKeyF16
        Case Else:                              vbKCode = "": Exit Sub
    End Select


If ShiftCtrlAlt > 0 Then
    GoSub SendWithControl
Else
    GoSub SendWithoutControl
End If


If Len(ST$) > P2 Then
    'If there are more characters in the string, remove those keys sent and start over.
    ST$ = Mid$(ST$, P2 + 1)
    GoTo Start
End If
Exit Sub
End If
'Section to send a Control Key and a Character
Set1$ = ")!@#$%^&*(" 'Characters above the numbers requiring SHIFT KEY
Set2$ = "`-=[]\;',./" 'Other miscellaneous characters
Set3$ = "~_+{}|:" & Chr(34) & "<>?" 'Miscellaneous characters requiring SHIFT KEY


If ShiftCtrlAlt > 0 Then
'Handle the three key problem which uses parentheses


If InStr(ST$, "(") > 0 Then    'Strip the Parentheses.
    ST$ = Mid$(ST$, 1, 1) & Mid$(ST$, 3, InStr(ST$, ")") - 3)
End If
vbKCode = Asc(UCase(Mid$(ST$, 2, 1)))
'Check for characters 0 to 9, and A to Z. Scan codes same as ASCII


If (vbKCode >= 48 And vbKCode <= 57) Or (vbKCode >= 65 And vbKCode <= 90) Then

    If ShiftCtrlAlt = vbKeyShift Then        'Handle the problem of the CAPSLOCK
        If CapsLockState = False Then
            GoSub SendWithControl
        Else
            GoSub SendWithoutControl
        End If
    Else
        GoSub SendWithControl
    End If
Else
    'Set the scan code for each miscellaneous character


    If InStr(Set1$, Mid$(ST$, 2, 1)) > 0 Then
        vbKCode = 47 + InStr(Set1$, Mid$(ST$, 2, 1))
    ElseIf InStr(Set2$, Mid$(ST$, 2, 1)) > 0 Then
        vbKCode = Choose(InStr(Set2$, Mid$(ST$, 2, 1)), 192, 189, 187, 219, _
        221, 220, 186, 222, 188, 190, 191)
    ElseIf InStr(Set3$, Mid$(ST$, 2, 1)) > 0 Then
        vbKCode = Choose(InStr(Set3$, Mid$(ST$, i, 1)), 192, 189, 187, 219, _
        221, 220, 186, 222, 188, 190, 191)
    End If
End If
'If there are more characters to print, remove the control key
'and the first character and go to the n
'     ext section. No control characters
'processed beyond this point.


If Len(ST$) > 2 Then
    ST$ = Mid$(ST$, 3)
Else
    Exit Sub
End If
End If
'********* SEND CHARACTER STRING *******
'     ***
'Send all remaining characters as text,
'     including control type characters
'such as (+^%{[) etc.
ShiftCtrlAlt = vbKeyShift 'Prepare to send the SHIFT KEY when needed
'Set the scan code for each character an
'     d then send it


For i = 1 To Len(ST$)
vbKCode = Asc(UCase(Mid$(ST$, i, 1)))


If InStr(Set1$, Mid$(ST$, i, 1)) > 0 Then
    vbKCode = 47 + InStr(Set1$, Mid$(ST$, i, 1))
    GoSub SendWithControl
ElseIf InStr(Set2$, Mid$(ST$, i, 1)) > 0 Then
    vbKCode = Choose(InStr(Set2$, Mid$(ST$, i, 1)), 192, 189, 187, 219, 221, _
    220, 186, 222, 188, 190, 191)
    GoSub SendWithoutControl
ElseIf InStr(Set3$, Mid$(ST$, i, 1)) > 0 Then
    vbKCode = Choose(InStr(Set3$, Mid$(ST$, i, 1)), 192, 189, 187, 219, 221, _
    220, 186, 222, 188, 190, 191)
    GoSub SendWithControl
Else
    'Check to see if the character is upper
    '     or lower case and then
    'determine whether to send the SHIFT KEY
    '     based upon whether or not
    'the CAPSLOCK is set.


    If Asc(Mid$(ST$, i, 1)) > vbKCode Then 'If true character is to be lowercase
        If CapsLockState = False Then
            GoSub SendWithoutControl
        Else
            GoSub SendWithControl
        End If
    Else
        If CapsLockState = False Then
            GoSub SendWithControl
        Else
            GoSub SendWithoutControl
        End If
    End If
End If
Next i
Exit Sub
'API call to send a Control Code and a C
'     haracter
SendWithControl:
keybd_event ShiftCtrlAlt, 0, 0, 0 'Control Key Down
keybd_event vbKCode, 0, 0, 0 'Character Key Down
keybd_event ShiftCtrlAlt, 0, &H2, 0 'Control Key Up
keybd_event vbKCode, 0, &H2, 0 'Character Key Up
Return
'API call to send only one Character
SendWithoutControl:
keybd_event vbKCode, 0, 0, 0 'Character Key Down
keybd_event vbKCode, 0, &H2, 0 'Character Key Up
Return
End Sub

'Public Function AttachhWnds(SRC As Long, DEST As Long, Optional Attach As Boolean = True) As Boolean
'    AttachhWnds = AttachThreadInput(GetProcessID(SRC), GetProcessID(DEST), Attach) = 0
'End Function
'Public Function Attachto(Dest As Long, Optional Attach As Boolean = True) As Boolean
'    Attachto = AttachThreadInput(App.ThreadID, GetWindowThreadProcessId(Dest, 0), Attach) = 0
'End Function

Public Sub SendKey(hwnd As Long, vKey As VKeys, Optional Release As Boolean)
    SendMessageByString hwnd, IIf(Release, WM_KEYUP, WM_KEYDOWN), vKey, 0
    'Call SendMessage(Text&, WM_SETTEXT, Len(Text1.Text), ByVal Text1.Text)
End Sub

Public Function KeyCode(ByVal sChar As String) As KeyCodeConstants
Dim bNt As Boolean
Dim iKeyCode As Integer
Dim b() As Byte
Dim iKey As Integer
Dim vKey As KeyCodeConstants
Dim iShift As ShiftConstants

   ' Determine if we have Unicode support or not:
   bNt = ((GetVersion() And &H80000000) = 0)
   
   ' Get the keyboard scan code for the character:
   If (bNt) Then
      b = sChar
      CopyMemory iKey, b(0), 2
      iKeyCode = VkKeyScanW(iKey)
   Else
      b = StrConv(sChar, vbFromUnicode)
      iKeyCode = VkKeyScan(b(0))
   End If
   
   KeyCode = (iKeyCode And &HFF&)

End Function

Public Function sKeyname(vKey As Long) As String
Dim lScanCode As Long
Dim sBuf As String
Dim lSize As Long
Dim b() As Byte, bNt As Boolean
    
    bNt = ((GetVersion() And &H80000000) = 0)
   ' Translate the virtual-key code into a scan code.
   If (bNt) Then
      lScanCode = MapVirtualKeyW(vKey, 0)
   Else
      lScanCode = MapVirtualKey(vKey, 0)
   End If
   
   ' GetKeyNameText retrieves the name of a key (the scan code
   ' must be in bits 16-23):
   lScanCode = lScanCode * &H10000
   If (bNt) Then
      ReDim b(0 To 512) As Byte
      lSize = GetKeyNameTextW(lScanCode, VarPtr(b(0)), 256)
      If (lSize > 0) Then
         sBuf = b
         sKeyname = Left$(sBuf, lSize)
      End If
   Else
      sBuf = SPACE$(256)
      lSize = GetKeyNameText(lScanCode, sBuf, 256)
      sKeyname = Left$(sBuf, lSize)
   End If
End Function
Public Function GetKeyboardString(ByVal sChar As String, Optional ByRef vKey As KeyCodeConstants, Optional ByRef iShift As ShiftConstants) As String
Dim lScanCode As Long
Dim b() As Byte
Dim sRet As String
Dim sBuf As String
Dim lSize As Long
Dim bNt As Boolean
Dim iKeyCode As Integer

   ' Determine if we have Unicode support or not:
   bNt = ((GetVersion() And &H80000000) = 0)
   
   ' Get the keyboard scan code for the character:
   If (bNt) Then
      b = sChar
      CopyMemory vKey, b(0), 2
      iKeyCode = VkKeyScanW(vKey)
   Else
      b = StrConv(sChar, vbFromUnicode)
      iKeyCode = VkKeyScan(b(0))
   End If
   
   ' Split into shift and key portions:
   iShift = (iKeyCode And &HFF00&) \ &H100&
   vKey = iKeyCode And &HFF&

   ' Build the string for the return state:
   sRet = _
      IIf(iShift And vbAltMask, "Alt+", vbNullString) & _
      IIf(iShift And vbCtrlMask, "Ctrl+", vbNullString) & _
      IIf(iShift And vbShiftMask, "Shift+", vbNullString)
   
   ' Translate the virtual-key code into a scan code.
   If (bNt) Then
      lScanCode = MapVirtualKeyW(vKey, 0)
   Else
      lScanCode = MapVirtualKey(vKey, 0)
   End If
   
   ' GetKeyNameText retrieves the name of a key (the scan code
   ' must be in bits 16-23):
   lScanCode = lScanCode * &H10000
   If (bNt) Then
      ReDim b(0 To 512) As Byte
      lSize = GetKeyNameTextW(lScanCode, VarPtr(b(0)), 256)
      If (lSize > 0) Then
         sBuf = b
         sRet = sRet & Left$(sBuf, lSize)
      End If
   Else
      sBuf = SPACE$(256)
      lSize = GetKeyNameText(lScanCode, sBuf, 256)
      sRet = sRet & Left$(sBuf, lSize)
   End If
      
   GetKeyboardString = sRet
      
End Function


Private Function PressKey(Key As VKeys, Optional Release As Boolean)
    'method 1
    keybd_event Key, 0, IIf(Release, 2, 0), 0
    
    'method 2
    'If Release Then
    '    keybd_event Key, 0, KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0
    'Else
    '    keybd_event Key, 0, KEYEVENTF_EXTENDEDKEY, 0
    'End If
    
    'method 3
    'SendKey 3736114, Key, Release
    
    'method 4
    'If SKEY Is Nothing Then Set SKEY = New cSendKeys
    'SKEY.PressKey Key
End Function

Public Sub VKeyPress(Key As VKeys, Optional CTRL As Boolean, Optional ALT As Boolean, Optional hwnd As Long = -1, Optional DoDown As Boolean = True, Optional DoUp As Boolean = True, Optional StillDown As Boolean = False)
    If hwnd <> -1 Then
        SetForegroundWindow hwnd
        Do Until GetForegroundWindow = hwnd
            DoEvents
        Loop
    End If
    
    If DoDown Then
        If Not StillDown Then
            If CTRL Then PressKey VK_CONTROL
            If ALT Then PressKey VK_MENU
        End If
        PressKey Key
    End If
    If DoUp Then
        PressKey Key, True
        If CTRL Then PressKey VK_CONTROL, True
        If ALT Then PressKey VK_MENU, True
    End If
End Sub
Public Sub KeyPress(Key As String, Optional CTRL As Boolean, Optional ALT As Boolean, Optional hwnd As Long = -1)
    VKeyPress Asc(Key), CTRL, ALT, hwnd
End Sub







Public Function Mouse_Loc() As POINTAPI
    Dim temp As POINTAPI
    GetCursorPos temp
    Mouse_Loc = temp
End Function
Public Sub Mouse_MoveTo(ByVal X As Long, ByVal Y As Long, Optional Absolute As Boolean)
    Dim xl As Double, yl As Double, xMax As Long, yMax As Long
   ' Move the mouse:
    If Absolute Then
        ' mouse_event ABSOLUTE coords run from 0 to 65535:
        xMax = Screen.Width \ Screen.TwipsPerPixelX
        yMax = Screen.Height \ Screen.TwipsPerPixelY
        xl = X * 65535 / xMax
        yl = Y * 65535 / yMax
        mouse_event MOUSEEVENTF_MOVE Or MOUSEEVENTF_ABSOLUTE, xl, yl, 0, 0
    Else
        mouse_event MOUSEEVENTF_MOVE, X, Y, 0, 0
    End If
End Sub
Public Sub Mouse_Click(Optional ByVal eButton As MouseButtonConstants = vbLeftButton, Optional Down As Boolean = True, Optional Up As Boolean = True)
    Dim lFlagDown As Long, lFlagUp As Long
    Select Case eButton
        Case vbRightButton
            lFlagDown = MOUSEEVENTF_RIGHTDOWN
            lFlagUp = MOUSEEVENTF_RIGHTUP
        Case vbMiddleButton
            lFlagDown = MOUSEEVENTF_MIDDLEDOWN
            lFlagUp = MOUSEEVENTF_MIDDLEUP
        Case vbLeftButton
            lFlagDown = MOUSEEVENTF_LEFTDOWN
            lFlagUp = MOUSEEVENTF_LEFTUP
    End Select
    ' A click = down then up
    If Down Then mouse_event lFlagDown Or MOUSEEVENTF_ABSOLUTE, 0, 0, 0, 0
    If Up Then mouse_event lFlagUp Or MOUSEEVENTF_ABSOLUTE, 0, 0, 0, 0
End Sub
Public Sub Mouse_Scroll(Y As Long, Optional Horizontal As Boolean)
    Const Delta As Long = 120
    mouse_event IIf(Horizontal, MOUSEEVENTF_HWHEEL, MOUSEEVENTF_WHEEL), 0, 0, Y * Delta, 0
End Sub
