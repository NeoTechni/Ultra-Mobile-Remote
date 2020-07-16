Attribute VB_Name = "ProgramAPI"
Option Explicit

Public Const NoneSelected As String = "There are no selected items"
Public Enum OpToDo
    DoNothing
    DoDelete
End Enum
Public DoOp As OpToDo
Public Bookmarks() As String, BookmarkCount As Long, NeedsSaving As Boolean, preview As Long

Public Function sec2time(ByVal whattime As Long) As String
    On Error Resume Next
    If InStr(whattime, ".") > 0 Then whattime = Left(whattime, ".") - 1
    Const time_min As Long = 60, time_hour As Long = 3600
    Dim time_hours As Byte, time_minutes As Byte, time_seconds As Byte

    time_hours = whattime \ time_hour
    whattime = whattime Mod time_hour
    time_minutes = whattime \ time_min
    whattime = whattime Mod time_min
    time_seconds = whattime

    'If time_hours = 0 Then
    '    sec2time = Format(time_minutes, "#0") & ":" & Format(time_seconds, "00")
    'Else
        sec2time = Format(time_hours, "#0:") & Format(time_minutes, "00") & ":" & Format(time_seconds, "00")
    'End If
End Function

Public Function IsInIDE() As Boolean
    IsInIDE = App.LogMode = 0
End Function

Public Sub ResizeLCARs(Optional Name As String = "frmbottom")
    Dim temp As Long, Width As Long, Top As Long
    'If GroupList(5).Visible Then
        Width = DestWidth / 2 - 130
    
        temp = LCAR_FindLCAR(Name, , 1) 'LEFT
        LCAR_ButtonList(temp).Width = Width + 1
        
        temp = LCAR_FindLCAR(Name, , 7) 'DELETE
        LCAR_ButtonList(temp).Width = Width + 1
        
        temp = LCAR_FindLCAR(Name, , 2) 'RIGHT
        With LCAR_ButtonList(temp)
            .Width = Width
            .x = DestWidth / 2 + 1
        End With
        
        temp = LCAR_FindLCAR(Name, , 8) 'RIGHT
        With LCAR_ButtonList(temp)
            .Width = Width
            .x = DestWidth / 2 + 1
        End With
        
        If StrComp(Name, "frmbottom", vbTextCompare) = 0 Then
            ResizeLCARs "frmnumbottom"
        Else
            temp = LCAR_FindLCAR(Name, , 12) 'minus 1000
            LCAR_ButtonList(temp).Width = Width / 2 - 1
            For temp = temp + 1 To temp + 3
                LCAR_ButtonList(temp).Width = LCAR_ButtonList(temp - 1).Width
                LCAR_ButtonList(temp).x = LCAR_ButtonList(temp - 1).x + Width / 2 + 1
            Next
            LCAR_ButtonList(temp - 1).Width = LCAR_ButtonList(temp - 1).Width + 1
        End If
    IsClean = False
End Sub
Public Sub HideAllGroups(Optional Except As Long = -1)
    LCARlists(0).Visible = False
    LCARlists(2).Visible = False
    
    Dim temp As Long
    'GroupList(5).Visible = True
    For temp = 3 To GroupCount - 1
        GroupList(temp).Visible = (temp = Except)
    Next
    
    ResizeLCARs
    IsClean = False
End Sub

Public Sub HideGroup(ID As Long, Optional Visible As Boolean)
    GroupList(ID).Visible = Visible
    IsClean = False
End Sub

Public Sub HideAllLists(Optional Except As Long = -1)
    Dim temp As Long
    If KeyboardIsVisible Then HideKeyboard
    For temp = 0 To LCARListCount - 1
        LCARlists(temp).Visible = (Except = temp)
    Next
    IsClean = False
End Sub

Public Sub RefreshPreview(Optional ForceText As String)
    Dim tempstr As String, tempstr2 As String, temp3 As Long
                With LCAR_ButtonList(preview)
                    .IsClean = False
                    '.Visible = True
                    If Len(ForceText) > 0 Then
                        .Text = ForceText
                    Else
                    
                    Select Case LCARlists(0).SelectedItems
                        Case 0: .Text = NoneSelected
                        Case 1
                            If LCARlists(0).SelectedItem > -1 And LCARlists(0).SelectedItem < LCARlists(0).ListCount Then
                                tempstr = LCARlists(0).ListItems(LCARlists(0).SelectedItem).Text
                                tempstr2 = UCase(LCARlists(0).ListItems(LCARlists(0).SelectedItem).Side)
                                temp3 = LCARlists(0).TotalSize
                                'If Len(tempstr2) > 0 Then
                                    If Len(tempstr2) = Empty Then
                                        .Text = tempstr & vbNewLine & "File Folder"
                                    Else
                                        tempstr = tempstr & "." & LCase(tempstr2)
                                        '.Text = tempstr & vbNewLine & FileTypeName(tempstr2, , "The ""*"" extention has no association at this time") & vbNewLine & "This file occupies " & SizeToText(temp3, " Quads", " KiloQuads", " MegaQuads", " GigaQuads", 2)
                                    End If
                                'Else
                                '    .Text = tempstr
                                'End If
                                
                            End If
                        Case Else
                            .Text = "There are " & LCARlists(0).SelectedItems & " selected items occupying a total of " & SizeToText(LCARlists(0).TotalSize, " Quads", " KiloQuads", " MegaQuads", " GigaQuads", 2)
                    End Select
                    
                    End If
                End With
                LCAR_DrawLCARs
End Sub


Public Function IsFontInstalled(Name As String) As Boolean
    Dim temp As Long
    For temp = 0 To Screen.FontCount - 1
        If StrComp(Name, Screen.Fonts(temp), vbTextCompare) = 0 Then
            IsFontInstalled = True
            Exit For
        End If
    Next
End Function
