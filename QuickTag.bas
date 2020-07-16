Attribute VB_Name = "QuickTag"
Option Explicit

Public Function LoadFile(Filename As String) As String
    On Error Resume Next
    Dim intFile As Integer, temp As String, allfile As String
    allfile = Empty
    intFile = FileLen(Filename)
    If DIR(Filename) <> Empty And Right(Filename, 1) <> "\" And intFile > 0 Then
        intFile = FreeFile()
        Open Filename For Input As intFile
            Do Until EOF(intFile)
                Line Input #intFile, temp
                allfile = allfile & temp & vbNewLine
            Loop
        Close intFile
        LoadFile = Left(allfile, Len(allfile) - 1)
    End If
End Function
Public Function SaveFile(Filename As String, Contents As String) As Boolean
    On Error Resume Next
    Dim temp As Long
    temp = FreeFile
    Open Filename For Output As temp
        Print #temp, Contents
    Close temp
    SaveFile = True
End Function
'Check to see if text is a tag (use on a processed array)
Public Function QTAG_isTag(Text As String) As Boolean
    QTAG_isTag = Left(Text, 1) = "[" And Right(Text, 1) = "]"
End Function

Private Function AddtoArray(tempstr, Count As Long, entry As String) As Long
    Count = Count + 1
    ReDim Preserve tempstr(Count)
    tempstr(Count - 1) = entry
    AddtoArray = Count
End Function

'Splits Text and Tags up into an array
Public Function QTAG_Split(Text As String, tempstr, Optional LeftSide As String = "[", Optional RightSide As String = "]") As Long
    Dim Count As Long, temp As Long, temp2 As Long, Start As Long
    temp = InStr(Text, LeftSide)
    Start = 1
    Do Until temp = 0
        If temp - Start > 0 Then
            'Debug.Print "TEXT: " & Mid(Text, Start, temp - Start)
            AddtoArray tempstr, Count, Mid(Text, Start, temp - Start)
        End If
    
        temp2 = InStr(temp + 1, Text, RightSide)
        
        'Debug.Print "TAG: " & Mid(Text, temp, temp2 - temp + 1)
        AddtoArray tempstr, Count, Mid(Text, temp, temp2 - temp + 1)
        
        temp = InStr(temp2 + 1, Text, LeftSide)
        Start = temp2 + 1
        
    Loop
    If temp2 < Len(Text) Then
        'Debug.Print "TEXT: " & Right(Text, Len(Text) - temp2)
        AddtoArray tempstr, Count, Right(Text, Len(Text) - temp2)
    End If
    
    QTAG_Split = Count
End Function

Public Function GetStart(Text As String, Optional Start As Long = 1) As Long
    Dim temp As Long
    If Start > 0 Then
        temp = InStr(Start, Text, " ") + 1
        If temp >= Len(Text) Then temp = 0
        GetStart = temp
    End If
End Function
Public Function GetEnd(Text As String, Start As Long) As Long
    Dim temp As Long, temp2 As Long, doit As Boolean
    For temp = Start To Len(Text)
        Select Case Mid(Text, temp, 1)
            Case "="
                Select Case Mid(Text, temp + 1, 1)
                    Case Is = "'"
                        temp2 = InStr(temp + 2, Text, "'")
                    Case Is = """"
                        temp2 = InStr(temp + 2, Text, """")
                    Case Else
                        temp2 = InStr(temp, Text, " ")
                End Select
                If temp2 = 0 Then temp2 = Len(Text)
                GetEnd = temp2
                Exit For
            Case " ", "]"
                GetEnd = temp - 1
                Exit For
        End Select
    Next
    If temp = Len(Text) Then GetEnd = Len(Text)
End Function
Private Function hasvalue(Text As String, Start As Long) As Boolean
    Dim temp As Long
    For temp = Start To Len(Text)
        Select Case Mid(Text, temp, 1)
            Case "="
                hasvalue = True
                Exit For
            Case " ", "]"
                Exit For
        End Select
    Next
End Function

Public Function QTAG_SplitValues(ByVal Tag As String, tempstr) As Long
    Tag = Mid(Tag, 2, Len(Tag) - 2)
    Dim temp As Long, temp2 As Long, Count As Long
    temp = GetStart(Tag)
    Do Until temp = 0 Or temp2 = Len(Tag)
        temp2 = GetEnd(Tag, temp)
        If temp2 = 0 Then temp2 = Len(Tag)
        AddtoArray tempstr, Count, Mid(Tag, temp, temp2 - temp + 1)
        temp = GetStart(Tag, temp2)
    Loop
    QTAG_SplitValues = Count
End Function

Public Function QTAG_GetValue(Tag As String, Name As String, Optional Default As String) As String
    Dim tempstr() As String, Count As Long, temp As Long
    Count = QTAG_SplitValues(Tag, tempstr)
    QTAG_GetValue = Default
    For temp = 0 To Count - 1
        If StrComp(QTAG_Name(tempstr(temp)), Name, vbTextCompare) = 0 Then
            QTAG_GetValue = QTAG_Value(tempstr(temp))
            Exit For
        End If
    Next
End Function

Public Function QTAG_ValueExists(Tag As String, Name As String) As Boolean
    Dim tempstr() As String, Count As Long, temp As Long
    Count = QTAG_SplitValues(Tag, tempstr)
    For temp = 0 To Count - 1
        If StrComp(QTAG_Name(tempstr(temp)), Name, vbTextCompare) = 0 Then
            QTAG_ValueExists = True
            Exit For
        End If
    Next
End Function

Public Function QTAG_Name(ByVal Text As String) As String
    If QTAG_isTag(Text) Then Text = Mid(Text, 2, Len(Text) - 2)
    If InStr(Text, "=") > 0 Then Text = Left(Text, InStr(Text, "=") - 1)
    If InStr(Text, " ") > 0 Then Text = Left(Text, InStr(Text, " ") - 1)
    QTAG_Name = Text
End Function

Private Function QTAG_Value(Text As String, Optional Clean As Boolean = True) As String
    Dim temp As String
    If InStr(Text, "=") > 0 Then
        temp = Right(Text, Len(Text) - InStr(Text, "="))
        If Clean Then
            If Left(temp, 1) = """" And Right(temp, 1) = """" Then temp = Mid(temp, 2, Len(temp) - 2)
            If Left(temp, 1) = "'" And Right(temp, 1) = "'" Then temp = Mid(temp, 2, Len(temp) - 2)
            temp = Replace(temp, """""", """")
        End If
        QTAG_Value = temp
    End If
End Function

'Works with embedded Quicktags (%tag=value%)
Public Function QTAG_SplitQ(ByVal Text As String, tempstr) As Long
    Dim temp As Long, Count As Long, Start As Long, tempstr2 As String
    Start = 1
    Text = Text & " "
    Do Until Start >= Len(Text)
        tempstr2 = QTAG_GrabBit(Text, Start)
        Start = Start + Len(tempstr2)
        AddtoArray tempstr, Count, tempstr2
    Loop
    QTAG_SplitQ = Count
End Function
Private Function QTAG_GrabBit(Text As String, Optional Start As Long) As String
    Dim temp As Long
    temp = InStr(Start + 1, Text, "%")
    If temp = 0 Then temp = Len(Text)
    
    If Mid(Text, Start, 1) = "%" Then
        QTAG_GrabBit = Mid(Text, Start, temp - Start + 1)
    Else
        QTAG_GrabBit = Mid(Text, Start, temp - Start)
    End If
End Function
Public Function QTAG_isEmbedded(Text As String) As Boolean
    QTAG_isEmbedded = Left(Text, 1) = "%" And Right(Text, 1) = "%"
End Function

'Converted to work with embedded and stand-alone tag types
Public Function QTAG_Side(ByVal Text As String, Optional LeftSide As Boolean = True, Optional Default As String) As String
    If QTAG_isEmbedded(Text) Or QTAG_isTag(Text) Then Text = Mid(Text, 2, Len(Text) - 2)
    
    If LeftSide Then
        If InStr(Text, " ") > 0 Then Text = Left(Text, InStr(Text, " ") - 1)
        If InStr(Text, "=") = 0 Then
            QTAG_Side = Trim(Text)
        Else
            QTAG_Side = Trim(Left(Text, InStr(Text, "=") - 1))
        End If
    Else
        If InStr(Text, "=") = 0 Then
            QTAG_Side = Trim(Default)
        Else
            QTAG_Side = Trim(Right(Text, Len(Text) - InStr(Text, "=")))
        End If
    End If
End Function

'Converted to work with embedded and stand-alone tag types
Public Function QTAG_TagsExist(tempstr() As String, ParamArray Tags() As Variant) As Boolean
    Dim temp As Long, temp2 As Long, tempstr2 As String, Found As Boolean
    For temp = 0 To UBound(tempstr)
        If QTAG_isEmbedded(tempstr(temp)) Or QTAG_isTag(tempstr(temp)) Then
            tempstr2 = QTAG_Side(tempstr(temp))
            For temp2 = 0 To UBound(Tags)
                If StrComp(tempstr2, CStr(Tags(temp2)), vbTextCompare) = 0 Then
                    Found = True
                    Exit For
                End If
            Next
            If Found Then Exit For
        End If
    Next
    QTAG_TagsExist = Found
End Function

'Converted to work with embedded and stand-alone tag types
Public Function QTAG_TagExists(tempstr() As String, Tag As String, Optional isEmbedded As Boolean = True) As Boolean
    QTAG_TagExists = QTAG_FindTag(tempstr, Tag, 0, isEmbedded) > -1
End Function '

'Converted to work with embedded and stand-alone tag types
Public Function QTAG_FindTag(tempstr() As String, Tag As String, Optional Start As Long, Optional isEmbedded As Boolean = True) As Long
    Dim temp As Long, doit As Boolean
    QTAG_FindTag = -1
    For temp = Start To UBound(tempstr)
        If isEmbedded Then
            doit = QTAG_isEmbedded(tempstr(temp))
        Else
            doit = QTAG_isTag(tempstr(temp))
        End If
    
        If doit Then
            If StrComp(Tag, QTAG_Side(tempstr(temp)), vbTextCompare) = 0 Then
                QTAG_FindTag = temp
                Exit For
            End If
        End If
    Next
End Function
Public Function QTAG_GetTagValue(tempstr() As String, Tag As String, Optional Index As Long = 1, Optional Default As String) As String
    Dim temp As Long, Count As Long 'QTAG_Split
    QTAG_GetTagValue = Default
    For temp = 0 To UBound(tempstr)
        If QTAG_isEmbedded(tempstr(temp)) Then
            If StrComp(Tag, QTAG_Side(tempstr(temp)), vbTextCompare) = 0 Then
                Count = Count + 1
                If Count = Index Then
                    QTAG_GetTagValue = QTAG_Side(tempstr(temp), False, Default)
                    Exit For
                End If
            End If
        End If
    Next
End Function

Public Sub SortAlphabetic(tempstr() As String, Start As Long, Finish As Long)
    Dim temp As Long, temp2 As Long, tempstr2 As String
    For temp = Start + 1 To Finish
         For temp2 = temp To Start + 1 Step -1
            If StrComp(tempstr(temp2), tempstr(temp2 - 1), vbTextCompare) = -1 Then
                tempstr2 = tempstr(temp2 - 1)
                tempstr(temp2 - 1) = tempstr(temp2)
                tempstr(temp2) = tempstr2
            Else
                Exit For
            End If
         Next
    Next
End Sub

'Debugging
Public Sub QTAG_Test()
    Dim tempstr() As String, Count As Long, temp As Long
    
    Count = QTAG_Split("<TEST>[TEST]<123>[123]abc", tempstr)
    For temp = 0 To Count - 1
        Debug.Print tempstr(temp)
    Next
    
    'Test for split tags/text
    'Count = QTAG_Split(Test, tempstr)
    
    Count = QTAG_SplitValues("[This is=""a test"" of='the' emergency]", tempstr)
    
    For temp = 0 To Count - 1
        Debug.Print tempstr(temp)
    Next
    
    Count = QTAG_SplitQ("%test%ing%test2%ation%test3%magic", tempstr)
    
    For temp = 0 To Count - 1
        Debug.Print tempstr(temp)
    Next
    
    
End Sub
