Attribute VB_Name = "Module1"
Option Explicit

Public Const LB_SETTABSTOPS = &H192
Public Const MAX_PATH = 260

Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal Hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal Hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Declare Function GetWindowThreadProcessId Lib "user32" (ByVal Hwnd As Long, lpdwProcessId As Long) As Long
Declare Function SendMessageArray Lib "user32" Alias "SendMessageA" (ByVal Hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Type Window
    Hwnd As Long
    ClassName As String
    ProcessID As Long
    ThreadID As Long
    WindowName As String
End Type
Public WindowList() As Window, WindowCount As Long

Private Function fEnumWindowsCallBack(ByVal Hwnd As Long, ByVal lpData As Long) As Long
Dim lResult    As Long
Dim lThreadId  As Long
Dim lProcessId As Long
Dim sWndName   As String
Dim sClassName As String
'
' This callback function is called by Windows (from the EnumWindows
' API call) for EVERY window that exists.  It populates the aWindowList
' array with a list of windows that we are interested in.
'
fEnumWindowsCallBack = 1
sClassName = Space$(MAX_PATH)
sWndName = Space$(MAX_PATH)

lResult = GetClassName(Hwnd, sClassName, MAX_PATH)
sClassName = Left$(sClassName, lResult)
lResult = GetWindowText(Hwnd, sWndName, MAX_PATH)
sWndName = Left$(sWndName, lResult)

lThreadId = GetWindowThreadProcessId(Hwnd, lProcessId)

AddWindow Hwnd, sClassName, lProcessId, lThreadId, sWndName
End Function

Public Function EnumWindows() As Boolean
    Dim Hwnd As Long
    ClearWindowList
    Call EnumWindows(AddressOf fEnumWindowsCallBack, Hwnd)
End Function

Public Sub ClearWindowList()
    WindowCount = 0
    ReDim WindowList(0)
End Sub
Private Function AddWindow(Hwnd As Long, ClassName As String, ProcessID As Long, ThreadID As Long, WindowName As String) As Long
    AddWindow = WindowCount
    WindowCount = WindowCount + 1
    ReDim Preserve WindowList(WindowCount)
    With WindowList(WindowCount - 1)
        .Hwnd = Hwnd
        .ClassName = ClassName
        .ProcessID = ProcessID
        .ThreadID = ThreadID
        .WindowName = WindowName
    End With
    Debug.Print "Hwnd: " & Hwnd & " Classname: " & ClassName & " ProcessID: " & ProcessID & " ThreadID: " & ThreadID & " WindowName: " & WindowName
End Function

