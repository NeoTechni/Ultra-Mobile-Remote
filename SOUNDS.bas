Attribute VB_Name = "ResourceSounds"
Option Explicit
'blips 101 102 103 106 107 108 109
'warning 104
'error 105
'unable to comply 110

Public Mute As Boolean, EmergencyDisable As Boolean

Public Const SND_SYNC = &H0
Public Const SND_ASYNC     As Long = &H1
Public Const SND_MEMORY    As Long = &H4
Public Const SND_NODEFAULT = &H2
Public Const SND_LOOP = &H8
Public Const SND_NOSTOP = &H10
Public Const Flags& = SND_ASYNC Or SND_NODEFAULT Or SND_MEMORY

Private Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long

Public Sub PlayRESSound(Index As Integer)
    On Error Resume Next
    If Mute Or EmergencyDisable Then Exit Sub
    Dim lFlags&, vAddress$
    lFlags = SND_ASYNC Or SND_MEMORY
    If Index > 100 Then
        vAddress = StrConv(LoadResData(Index, "CUSTOM"), vbUnicode)
        sndPlaySound vAddress, lFlags
    End If
End Sub

Public Sub PlayRandomSound(ParamArray Choices() As Variant)
    Dim temp As Integer
    Randomize Timer
    temp = Rnd * UBound(Choices)
    PlayRESSound Val(Choices(temp))
End Sub
