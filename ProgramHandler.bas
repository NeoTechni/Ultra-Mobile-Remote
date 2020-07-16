Attribute VB_Name = "ProgByHWND"
Option Explicit

Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Declare Function SendMessageArray Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Type Window
    hwnd As Long
    ClassName As String
    ProcessID As Long
    ThreadID As Long
    WindowName As String
End Type
Public WindowList() As Window, WindowCount As Long

Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long

Public Declare Function GetForegroundWindow Lib "user32" () As Long

Private Declare Function OpenProcess Lib "Kernel32.dll" (ByVal dwDesiredAccessas As Long, ByVal bInheritHandle As Long, ByVal dwProcId As Long) As Long
Private Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cB As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameEx Lib "psapi" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpFilename As String, ByVal nSize As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long

Private Const MAX_PATH As Integer = 260
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16
Private Const LB_SETTABSTOPS = &H192

Private Type PROCESSENTRY32
   dwSize As Long
   cntUsage As Long
   th32ProcessID As Long
   th32DefaultHeapID As Long
   th32ModuleID As Long
   cntThreads As Long
   th32ParentProcessID As Long
   pcPriClassBase As Long
   dwFlags As Long
   szExeFile As String * MAX_PATH
End Type

Private OnlyVis As Boolean, Pattern As String

Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long

Private Const GW_HWNDNEXT = 2
Private Const GW_CHILD = 5

Public Function ThreadID(hwnd As Long) As Long
    Dim GetPID As Long
    GetWindowThreadProcessId hwnd, GetPID
    ThreadID = GetPID
End Function

Public Function FindWindowLike(strPartOfCaption As String) As Long
    Dim hwnd As Long
    Dim strCurrentWindowText As String
    Dim R As Integer
    FindWindowLike = -1
    
    hwnd = GetForegroundWindow
    Do Until hwnd = 0
        strCurrentWindowText = SPACE$(255)
        R = GetWindowText(hwnd, strCurrentWindowText, 255)
        strCurrentWindowText = Left$(strCurrentWindowText, R)
        'hWnd = GetWindow(hWnd, GW_CHILD)
        If InStr(1, strCurrentWindowText, strPartOfCaption, vbTextCompare) > 0 Then ' GoTo Found
            FindWindowLike = hwnd
            hwnd = 0
        Else
            'Debug.Print strCurrentWindowText
            hwnd = GetWindow(hwnd, GW_HWNDNEXT)
        End If
    Loop
End Function

'returns the path to the EXE of HWND, if no HWND is given it uses the currently in-focus window
Public Function GetHwndEXE(Optional ByVal hwnd As Long = -1) As String
    Dim lProcessId As Long, lThread As Long
    Dim lProcessHandle As Long
    Dim sName As String, lModule As Long
    Dim bMore As Boolean, tPROCESS As PROCESSENTRY32
    Dim lSnapShot As Long, tName As String
    If hwnd = -1 Then hwnd = GetForegroundWindow
   
    lThread = GetWindowThreadProcessId(hwnd, lProcessId)
   
        'Win NT
        'Create an Instance of the Process
        lProcessHandle = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, 0&, lProcessId)
       
        'If the Process was successfully created, get the EXE
        If lProcessHandle Then
            'Just get the First Module, all we need is the Handle to get the Filename..
            If EnumProcessModules(lProcessHandle, lModule, 4, 0&) Then
            
                sName = SPACE(MAX_PATH)
                Call GetModuleFileNameExA(lProcessHandle, lModule, sName, Len(sName))
                tName = RTrim(sName)
                Do While Right(tName, 1) = " " Or Asc(Right(tName, 1)) = 0
                    tName = Left(tName, Len(tName) - 1)
                Loop
                GetHwndEXE = tName
                'If GetPath Then Path = ExePathFromHwndPSAPI(Hwnd)
            End If
            'Close the Process Handle
            Call CloseHandle(lProcessHandle)
        End If
    'End If
End Function

Private Function fEnumWindowsCallBack(ByVal hwnd As Long, ByVal lpData As Long) As Long
    Dim lResult    As Long
    Dim lThreadId  As Long
    Dim lProcessId As Long
    Dim sWndName   As String
    Dim sClassName As String
    Dim Visible As Boolean, doAdd As Boolean
    ' This callback function is called by Windows (from the EnumWindows
    ' API call) for EVERY window that exists.  It populates the aWindowList
    ' array with a list of windows that we are interested in.

    Visible = True
    If OnlyVis Then Visible = IsWindowVisible(hwnd)

    fEnumWindowsCallBack = 1

    If Visible Then
        sClassName = SPACE$(MAX_PATH)
        sWndName = SPACE$(MAX_PATH)

        lResult = GetClassName(hwnd, sClassName, MAX_PATH)
        sClassName = Left$(sClassName, lResult)
        lResult = GetWindowText(hwnd, sWndName, MAX_PATH)
        sWndName = Left$(sWndName, lResult)

        lThreadId = GetWindowThreadProcessId(hwnd, lProcessId)

        If OnlyVis Then Visible = Len(sWndName) > 0
        If Visible Then
            If Len(Pattern) > 0 Then
                If IsPattern(Pattern) Then
                    doAdd = islike(sWndName, Pattern)
                Else
                    doAdd = InStr(1, sWndName, Pattern, vbTextCompare) > 0
                End If
            Else
                doAdd = True
            End If
            If doAdd Then AddWindow hwnd, sClassName, lProcessId, lThreadId, sWndName
        End If
    End If
End Function

Public Function IsPattern(Text As String) As Boolean
    IsPattern = InStr(Text, "*") > 0 Or InStr(Text, "?") > 0
End Function


Public Function fEnumWindows(Optional OnlyVisible As Boolean, Optional filter As String) As Long
    Dim hwnd As Long
    ClearWindowList
    OnlyVis = OnlyVisible 'parameters cannot be passed to the callback
    Pattern = filter
    Call EnumWindows(AddressOf fEnumWindowsCallBack, hwnd)
    fEnumWindows = WindowCount
End Function

Public Sub ClearWindowList()
    WindowCount = 0
    ReDim WindowList(0)
End Sub

Private Function AddWindow(hwnd As Long, ClassName As String, ProcessID As Long, ThreadID As Long, WindowName As String) As Long
    AddWindow = WindowCount
    WindowCount = WindowCount + 1
    ReDim Preserve WindowList(WindowCount)
    With WindowList(WindowCount - 1)
        .hwnd = hwnd
        .ClassName = ClassName
        .ProcessID = ProcessID
        .ThreadID = ThreadID
        .WindowName = WindowName
    End With
    'Debug.Print "Hwnd: " & hwnd & " Classname: " & ClassName & " ProcessID: " & ProcessID & " ThreadID: " & ThreadID & " WindowName: " & WindowName
End Function

Public Function WindowTitle(ByVal lHwnd As Long) As String
Dim lLen As Long
Dim sBuf As String

    ' Get the Window Title:
    lLen = GetWindowTextLength(lHwnd)
    If (lLen > 0) Then
        sBuf = String$(lLen + 1, 0)
        lLen = GetWindowText(lHwnd, sBuf, lLen + 1)
        WindowTitle = Left$(sBuf, lLen)
    End If
    
End Function

Public Function ClassName(ByVal lHwnd As Long) As String
Dim lLen As Long
Dim sBuf As String
    lLen = 260
    sBuf = String$(lLen, 0)
    lLen = GetClassName(lHwnd, sBuf, lLen)
    If (lLen <> 0) Then
        ClassName = Left$(sBuf, lLen)
    End If
End Function

Public Sub ActivateWindow(ByVal lHwnd As Long)
    SetForegroundWindow lHwnd
End Sub


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

