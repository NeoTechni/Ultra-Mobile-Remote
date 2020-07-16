VERSION 5.00
Begin VB.UserControl Hini 
   ClientHeight    =   480
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   480
   InvisibleAtRuntime=   -1  'True
   Picture         =   "Hini4.0.ctx":0000
   ScaleHeight     =   480
   ScaleWidth      =   480
   ToolboxBitmap   =   "Hini4.0.ctx":0C42
End
Attribute VB_Name = "Hini"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Hierarchical INI format reading and editing functions
'Backwards compatable with standard INI and XML files

'The Hini file format is similar to ini files as it contains groups of key=value
'pairs within [sections]. [section]s in hini files can now contain more [section]s
'which in turn can contain more key=value pairs and more [section]s and so forth
'a section is ended with a [/], known as a root, everything after that in no longer in the
'current [section]. Unlike ini files, the root itself can contain key=value pairs
'if a comment contains a '=' then it will be treated as a key, but this shouldnt matter much
'Can now have multiple keys using the same name as it can now check by instance/index

'Any time index is a parameter, and optional with a default of 0, its referring to the multikey instance
'if given 0 or 1, it'll use the first key with that name it encounters, and 2 is second, and so on

'Function name          Parameters                  Description

'The following are file manipulation commands
'SaveFile               filename                    Saves the loaded/created hini file to filename
'Loadfile               filename                    Loads a hierarchical ini file
'CloseFile                                          purges the currently loaded hini file
'LoadOldFile            filename                    Load an old/standard ini file as a hinifile, wont be converted to a hierarchy or multikey, [/] section will become root
'SaveOldFile            filename                    Saves a hini file as an old/standard ini file, root keys will be put in the [/] section, multikeys will have their instance added in brackets
'LoadXMLfile            filename                    Loads an XML file as a hini file
'LoadWholeFile          filename                    Loads any file and returns the result as a string

'The following functions are the userfriendly versions of commands listed below
'SetKey                 Section, Key, Value         Set the value of the key in the section to value
'GetKey                 Section, Key, [Default]     Gets the contents of the key in the section, returns [default] if it doesnt exists
'CountKeys              Section                     Gets the number of keys in Section
'CountSections          Section                     Gets the number of sections in Section
'RenameKey              Section, Key, Name          Sets the name of key in section to Name
'RenameSection          Section, Name               Sets the name of Section to Name (do not include the path to the section in the new name)
'EnumSections           Section, Array,[Key, Index] Fills the array with: a list of sections within the section specified
'                                                   If a Key(Index) is given, it returns that Key=Value pair from each sub-section as well
'EnumSectionsLike       Section, Array, filter      Fills the array with: a list of sections matching the filter provided within the section specified
'EnumKeys               Section, Array, [recurse]   Fills the array with: a list of key names, a list of key values in the section specified
'                                                   If recursive is true, it returns keys within all sub-sections as well
'EnumMultiKey           Section, Key, Array         Fills the array with: a list of the contents of each key named key within the section specified
'KeyExists              Section, Key                Returns whether or not the key in the section exists
'SectionExists          Section                     Returns whether or not the section exists
'CreateSection          Section                     Creates every level of section in the section path if they dont exist
'Deletesection          Section                     Deletes section and all its contents
'Deletekey              Section, Key                Deletes the key within the section

'GetKeyAtIndex
'GetMultiKeyIndex
'GetKeysIndex           Section, Key                Returns the keys index number in the section
'SectionAtIndex         Section, Index              Returns the section with at index inside the section specified
'GetSectionContents
'SetSectionContents
'AddSectionContents

'AddMultiKey
'CountMultiKeys
'DeleteMultiKey
'SetMultiKey
'Section2String

'islike
'LineCount                                           Returns the number of loaded line entries
'RemoveDoubles                                       Replaces duplicate section names with Name(###)

Private Type entry
    level As Long
    contents As String
End Type

Private Enum errcode 'wow did this get not used fast
    err_none = 0
    err_filenotloaded = 1
    err_filedoesntexist = 2
    err_HINI_sectionexists = 3
    err_sectiondoesntexist = 4
    err_HINI_keyexists = 5
    err_keydoesntexist = 6
    err_sectionhasnosections = 7
    err_sectionhasnokeys = 8
    err_filenotsaved = 9
End Enum

Private Enum KeyType
    key_Comment
    key_Section
    key_Root
    key_KeyValuePair
End Enum

Dim stag As String, inifile() As entry, entrycount As Long, ErrorCode As errcode, isloaded As Boolean

Public Event XMLProgress(Position As Long, Finish As Long) 'Used by the obsolete XML loading code, not used in the LoadXML

Public Sub AddMultiKey(Section As String, Key As String, Optional Value As String)
    CreateKey Section, Key, Value, CountMultiKeys(Section, Key) + 1
End Sub

Public Sub AddSectionContents(ByVal Section As String, Text As String, Optional IncludeTags As Boolean)
    Dim tempstr() As String, temp As Long, Currlevel As Long, temp2 As Long
    For temp = 0 To UBound(tempstr)
        If (IncludeTags And temp = 0) Or (IncludeTags And temp = UBound(tempstr)) Then temp = temp + 1
            Select Case TextType(tempstr(temp))
                Case key_Comment, key_KeyValuePair 'Treat keys as comments to support multikey!
                    temp2 = qualifiedsectionhandle(Section)
                    If temp2 > -1 Then
                        Currlevel = inifile(Currlevel).level - 1
                        temp2 = findroot(Currlevel)
                        insert temp2, Currlevel, tempstr(temp)
                    End If
                Case key_Section
                    Section = ChkSection(Section, stripsection(tempstr(temp)))
                Case key_Root '[/]
                    If InStr(1, Section, "\") = 0 Then
                        Section = Empty 'Only 1 section, blank it out
                    Else
                        Section = Left(Section, InStrRev(Section, "\") - 1) 'More than 1, remove the last one
                    End If
            End Select
    Next
End Sub

Public Sub CreateSection(Section As String)
If Len(Section) = 0 Then Exit Sub
Dim tempstr() As String, tempstr2() As Long, count As Long
tempstr = Split(Section, "\")
ReDim Preserve tempstr2(UBound(tempstr)) As Long
For count = 0 To UBound(tempstr)
    If count = 0 Then
        If HINI_sectionexists(0, tempstr(0)) = False Then
            newrootsection tempstr(0)
        End If
        tempstr2(0) = handlerootsections(tempstr(0))
    Else
        If HINI_sectionexists(tempstr2(count - 1), tempstr(count)) = False Then
            HINI_createsection tempstr2(count - 1), tempstr(count)
        End If
        tempstr2(count) = sectionhandle(tempstr2(count - 1), tempstr(count))
    End If
Next
End Sub

'Erases the currently loaded data
Public Sub CloseFile()
    entrycount = 0
    ReDim inifile(0)
    isloaded = False
End Sub

'Gets the count of key=value pairs within a section
Public Function CountKeys(Section As String) As Long
    If entrycount = 0 Then Exit Function
    Dim temp As Long
    temp = qualifiedsectionhandle(Section)
    If temp > 0 Then
        CountKeys = HINI_countkeys(temp)
    Else
        CountKeys = countrootkeys
    End If
End Function

Public Function CountMultiKeys(Optional Section As String, Optional Key As String) As Long
    If entrycount = 0 Then Exit Function
    CountMultiKeys = HINI_countmultikey(qualifiedsectionhandle(Section), Key)
End Function
'Gets the count of sub-sections within a section
Public Function CountSections(Optional Section As String)
    If entrycount = 0 Then Exit Function
    If IsMissing(Section) Then
        CountSections = CountRootSections
    Else
        Dim temp As Long
        temp = qualifiedsectionhandle(Section)
        If temp > 0 Then
            CountSections = HINI_countsections(temp)
        End If
    End If
End Function

Public Sub DeleteKey(Section As String, Key As String, Optional Index As Long = 0)
 If entrycount = 0 Then Exit Sub
       Dim temp As Long
    temp = qualifiedsectionhandle(Section)
    If temp > 0 Then
        temp = keyhandle(temp, Key, Index)
        removerange temp, temp
    End If
End Sub

Public Sub DeleteMultiKey(Section As String, Key As String)
    If entrycount = 0 Then Exit Sub
    Dim Start As Long, count As Long, temp As Long, count2 As Long, count3 As Long
    Start = qualifiedsectionhandle(Section)
    If Start = 0 Then
        temp = 0
        Start = 1
    Else
        temp = inifile(Start).level + 1
    End If
    count3 = HINI_countmultikey(Start, Key)
    If count3 > 0 Then
        For count = findroot(Start) To Start Step -1
            If inifile(count).level = temp Then
                If isvalue(inifile(count).contents) Then
                    If StrComp(stripname(inifile(count).contents), Key, vbTextCompare) = 0 Then
                        removerange count, count
                    End If
                End If
            End If
        Next
    End If
End Sub

Public Sub DeleteSection(Section As String, Optional LeaveTags As Boolean)
    If entrycount = 0 Then Exit Sub
    Dim temp As Long
    temp = qualifiedsectionhandle(Section)
    If temp > 0 Then
        removesection temp
    End If
End Sub

Public Function EnumKeys(Section As String, strArray, Optional Recursive As Boolean) As Long
    Dim temp As Long
    If entrycount = 0 Then Exit Function
    If Len(Section) > 0 Then
        temp = qualifiedsectionhandle(Section)
        If temp > 0 Then
            EnumKeys = HINI_EnumKeys(temp, strArray, Recursive)
        End If
    Else
        EnumKeys = EnumRootKeys(strArray)
    End If
End Function

Public Function EnumMultiKey(Section As String, Key As String, strArray) As Long
    If entrycount = 0 Then Exit Function
        Dim Start As Long, count As Long, temp As Long, count2 As Long, count3 As Long
    Start = qualifiedsectionhandle(Section)
    If Start = 0 Then
        temp = 0
        Start = 1
    Else
        temp = inifile(Start).level + 1
    End If

    count3 = HINI_countmultikey(Start, Key)
    If count3 > 0 Then
    ReDim Preserve strArray(1 To count3) As String
    For count = Start To findroot(Start)
        If inifile(count).level = temp Then
            If isvalue(inifile(count).contents) Then
                If StrComp(stripname(inifile(count).contents), Key, vbTextCompare) = 0 Then
                    count2 = count2 + 1
                    strArray(count2) = stripvalue(inifile(count).contents)
                End If
            End If
        End If
    Next
    End If
    EnumMultiKey = count3
End Function

Public Function EnumMultiSections(ByVal Section As String, strArray) As Long
    Dim temp As Long, temparray() As String, count As Long, temp2 As Long, temp3 As Long, name As String, Doit As Boolean
    'temp        = qualifiedsectionhandle to section
    'temparray() = enumerated section list in SECTION
    'count       = current item being searched
    'strarray()  = search results
    'temp2       = section count of SECTION
    'temp3       = search result count
    
    If entrycount > 0 Then
        If Len(Section) > 0 Then
            name = Right(Section, Len(Section) - InStrRev(Section, "\"))
            Section = Left(Section, InStrRev(Section, "\") - 1)
        
            temp = qualifiedsectionhandle(Section)
            If temp > 0 Then temp2 = EnumSections(Section, temparray)

        
            For count = 1 To temp2
                Doit = False
                If StrComp(Left(temparray(count), Len(name)), name, vbTextCompare) = 0 Then
                    If Len(temparray(count)) >= Len(name) + 3 Then
                        If Mid(temparray(count), Len(name), 2) = " (" Then
                            If Right(temparray(count), 1) = "(" Then
                                Doit = IsNumeric(Mid(temparray(count), Len(name) + 2, temparray(count) - 3))
                            End If
                        End If
                    Else
                        Doit = Len(temparray(count)) = Len(name)
                    End If
                End If
                If Doit Then
                    temp3 = temp3 + 1
                    ReDim Preserve strArray(1 To temp3)
                    strArray(temp3) = temparray(count)
                End If
            Next
            EnumMultiSections = temp3
        
        End If
    End If
End Function
'Enumerates all the sub-sections within a section, can grab one key=value pair as well
Public Function EnumSections(Section As String, strArray, Optional Key As String, Optional Index As Long) As Long
    Dim temp As Long
    If entrycount > 0 Then
        If Len(Section) > 0 Then
            temp = qualifiedsectionhandle(Section)
            If temp > 0 Then EnumSections = HINI_EnumSections(temp, strArray, Key, Index)
        Else
            EnumSections = EnumRootSections(strArray, Key, Index)
        End If
    End If
End Function

'Enumerates all the sub-sections within a section that have a name like the filter
Public Function EnumSectionsLike(Section As String, strArray, filter As String) As Long
    Dim temp As Long, temparray() As String, count As Long, temp2 As Long, temp3 As Long
    'temp        = qualifiedsectionhandle to section
    'temparray() = enumerated section list in SECTION
    'count       = current item being searched
    'strarray()  = search results
    'temp2       = section count of SECTION
    'temp3       = search result count
    
    If entrycount > 0 Then
        If Len(Section) > 0 Then
            temp = qualifiedsectionhandle(Section)
            If temp > 0 Then temp2 = HINI_EnumSections(temp, temparray)
        Else
            temp2 = EnumRootSections(temparray)
        End If
        For count = 1 To temp2
            If islike(filter, temparray(count)) Then
                temp3 = temp3 + 1
                ReDim Preserve strArray(1 To temp3)
                strArray(temp3) = temparray(count)
            End If
        Next
        EnumSectionsLike = temp3
    End If
End Function

'Gets the value of a key=value pair
Public Function GetKey(Section As String, Key As String, Optional Default As String, Optional Index As Long = 0) As String
    If entrycount = 0 Then Exit Function
    Dim temp As Long
    temp = qualifiedsectionhandle(Section)
    GetKey = Default
    If temp > 0 Then
        If HINI_keyexists(temp, Key) Then GetKey = HINI_GetKey(temp, Key, Index)
    End If
End Function

Public Function GetKeyAtIndex(Section As String, Index As Long, Optional Side As Long = 2) As String
    Dim temp As Long
    temp = qualifiedsectionhandle(Section)
    If temp > -1 Then GetKeyAtIndex = keyindex(temp, Index, Side)
End Function
Public Function GetKeysIndex(Section As String, Key As String, Optional Index As Long = 0) As Long
    If entrycount = 0 Then Exit Function
    Dim tempstr() As String, temp As Long, count As Long, count2 As Long
    temp = qualifiedsectionhandle(Section)
    If temp > 0 Then
        HINI_EnumKeys temp, tempstr
        For count = 1 To HINI_countkeys(temp)
            If StrComp(tempstr(count), Key, vbTextCompare) = 0 Then
                count2 = count2 + 1
                If count2 = Index Or Index = 0 Then
                    GetKeysIndex = count
                End If
            End If
        Next
    End If
End Function
Public Function GetMultiKeyIndex(Section As String, Key As String, Value As String, Optional CompareMethod As VbCompareMethod = vbTextCompare, Optional Start As Long = 1) As Long
    Dim strArray() As String, temp As Long, count As Long
    count = EnumMultiKey(Section, Key, strArray)
    GetMultiKeyIndex = -1
    For temp = Start To count
        If StrComp(strArray(temp), Value, CompareMethod) = 0 Then
            GetMultiKeyIndex = temp
            Exit For
        End If
    Next
End Function
Public Function GetSectionContents(Section As String, Optional IncludeTags As Boolean) As String
    If entrycount = 0 Then Exit Function
    Dim temp As Long, temp2 As Long, count As Long, tempstr As String, tabs As Long
    temp = qualifiedsectionhandle(Section)
    If temp > 0 Then
        temp2 = findroot(temp)
        If IncludeTags Then tempstr = inifile(temp).contents
        For count = temp + 1 To temp2 - 1
            If IncludeTags Then
                tabs = inifile(count).level - inifile(temp).level
            Else
                tabs = inifile(count).level - inifile(temp + 1).level
            End If
            If Len(tempstr) = 0 Then
                tempstr = String(tabs, vbTab) & inifile(count).contents
            Else
                tempstr = tempstr & vbNewLine & String(tabs, vbTab) & inifile(count).contents
            End If
        Next
        If IncludeTags Then tempstr = tempstr & vbNewLine & inifile(temp2).contents
    End If
    GetSectionContents = tempstr
End Function

Public Function KeyExists(Section As String, Key As String, Optional Index As Long = 0) As Boolean
    If entrycount = 0 Then Exit Function
    Dim temp As Long
    temp = qualifiedsectionhandle(Section)
    KeyExists = False
    If temp > 0 Then
        KeyExists = HINI_keyexists(temp, Key, Index)
    End If
End Function

Public Function islike(filter As String, expression As String) As Boolean 'islike("*.exe", "test.exe")
    Dim tempstr() As String, count As Long
    expression = LCase(expression)
    filter = LCase(filter)
    If InStr(filter, ";") > 0 Then
        tempstr = Split(filter, ";")
        For count = 0 To UBound(tempstr)
            'If Expression Like tempstr(count) Then IsLike = True
            If MiniLike(tempstr(count), expression) Then islike = True
        Next
    Else
        'IsLike = Expression Like Filter
        islike = MiniLike(filter, expression)
    End If
End Function

Public Function LineCount() As Long
    LineCount = entrycount
End Function
'Loads a HINI file
Public Function Loadfile(ByVal Filename As String, Optional appendtoroot As String = Empty) As Boolean
    On Error Resume Next
    If FileLen(Filename) = 0 Then Exit Function
    Dim tempfile As Long, Currlevel As Long, tempstr As String, continue As Boolean
    If Len(appendtoroot) = 0 Then
        entrycount = 0
        Currlevel = 0
    Else
        entrycount = entrycount + 1
        ReDim Preserve inifile(1 To entrycount)
        inifile(entrycount).level = 0
        inifile(entrycount).contents = "[" & appendtoroot & "]"
        Currlevel = 1
    End If
    tempfile = FreeFile
    ErrorCode = 0
    Filename = Replace(Filename, "\\", "\")

    If Dir(Filename) <> Empty Then
        Open Filename For Input As tempfile
            Do Until EOF(tempfile)
                Line Input #tempfile, tempstr
                Currlevel = HandleText(tempstr, Currlevel)
            Loop
        Close tempfile
        AccountForMissingRoots Currlevel
        isloaded = True
        Loadfile = True
    Else
        ErrorCode = err_filedoesntexist
    End If
End Function

'Loads an INI file
Public Function LoadOldFile(ByVal Filename As String) As Boolean
    On Error Resume Next
    If FileLen(Filename) = 0 Then Exit Function
    Dim tempfile As Long, Currlevel As Long, tempstr As String, continue As Boolean
    entrycount = 0
    Currlevel = 0
    tempfile = FreeFile

    isloaded = False
    Filename = Replace(Filename, "\\", "\")
    If Dir(Filename) <> Empty Then
        Open Filename For Input As tempfile
            Do Until EOF(tempfile)
                Line Input #tempfile, tempstr
                tempstr = Replace(tempstr, vbTab, Empty)
                If tempstr <> Empty And tempstr <> "[/]" Then 'removes blank lines and the [/] from hini2ini converted files
                    entrycount = entrycount + 1
                    If isSection(tempstr) And Currlevel > 0 Then
                        tempstr = Replace(tempstr, "\", "/")
                        entrycount = entrycount + 1
                        ReDim Preserve inifile(1 To entrycount)
                        inifile(entrycount - 1).level = 0
                        inifile(entrycount - 1).contents = "[/]"
                        Currlevel = Currlevel - 1
                    Else
                        If entrycount = 1 Then
                            ReDim inifile(1 To 1)
                        Else
                            ReDim Preserve inifile(1 To entrycount)
                        End If
                    End If
                    inifile(entrycount).level = Currlevel
                    inifile(entrycount).contents = tempstr
                    If isSection(tempstr) Then Currlevel = Currlevel + 1
                End If
            Loop
            If Currlevel > 0 Then
                entrycount = entrycount + 1
                ReDim Preserve inifile(1 To entrycount)
                inifile(entrycount).level = 0
                inifile(entrycount).contents = "[/]"
            End If
        Close tempfile
        isloaded = True
        LoadOldFile = True
    End If
End Function



'Loads the contents of any file and returns it as a string
Public Function LoadWholeFile(ByVal Filename As String) As String
    On Error Resume Next
    Filename = Replace(Filename, "\\", "\")
    If FileLen(Filename) = 0 Then Exit Function
    Dim temp As Long, tempstr As String, tempstr2 As String
    temp = FreeFile
    If Dir(Filename) <> Filename Then
        Open Filename For Input As temp
            Do Until EOF(temp)
                Line Input #temp, tempstr
                If tempstr2 <> Empty Then tempstr2 = tempstr2 & vbNewLine
                tempstr2 = tempstr2 & tempstr
                DoEvents
            Loop
            LoadWholeFile = tempstr2
        Close temp
    End If
End Function

Public Function LoadXMLfile(Filename As String, Optional Compress As Boolean = True) As Boolean
    Dim XML As String, tempstr As String, Currlevel As Long
    XML = LoadWholeFile(Filename)
    Do Until Len(XML) = 0
        tempstr = XML_Grabword(XML)
        XML = Right(XML, Len(XML) - Len(tempstr))
        tempstr = XML_Trim(tempstr)
        
        If Len(tempstr) = 0 Then
            XML = Trim(XML_Trim(XML))
        Else
            If Left(tempstr, 1) = "<" Then
                'isatag
                XML_ParseTag tempstr, Currlevel
            Else
                'istext
                ReDimPreserve entrycount, inifile, Currlevel, "Node=" & tempstr
            End If
        End If
    Loop
    isloaded = True
    RemoveDoubles
    If Compress Then XML_Compress
    AccountForMissingRoots Currlevel
    LoadXMLfile = True
End Function


'Sets the key name of a key=value pair
Public Function RenameKey(Section As String, Key As String, NewName As String) As Boolean
    If entrycount = 0 Then Exit Function
    Dim temp As Long
    temp = qualifiedsectionhandle(Section)
    If temp > 0 Then
        If HINI_keyexists(temp, Key) = True Then ' And HINI_keyexists(temp, NewName) = False Then
            HINI_RenameKey temp, Key, NewName
            RenameKey = True
        End If
    End If
End Function

'Sets the name of a section if the new name doesn't exist already
Public Function RenameSection(Section As String, name As String) As Boolean
    If entrycount = 0 Then Exit Function
    Dim temp As Long
    temp = qualifiedsectionhandle(Section)
    If temp > 0 Then
        If InStr(name, "\") > 0 Then name = Right(name, Len(name) - InStrRev(name, "\"))
        HINI_renamesection temp, name
        RenameSection = True
    End If
End Function

Public Sub RemoveDoubles()
    If entrycount = 0 Then Exit Sub
         Dim count As Long, count2 As Long
     For count = 1 To entrycount
        If isnotroot(inifile(count).contents) Then
            removedoublessection count
            'This code is by Techni Myoko, and anyone who claims otherwise is a liar
            'removedoublekey count ' not needed with multikey
        End If
     Next
End Sub

'Saves the currently loaded data as a HINI file
Public Function SaveFile(ByVal Filename As String) As Boolean
    On Error Resume Next
    If entrycount = 0 Then Exit Function
    Dim tempfile As Long, count As Long
    If isloaded Then
        tempfile = FreeFile
        Filename = Replace(Filename, "\\", "\")
        Open Filename For Output As tempfile
            For count = 1 To entrycount
                Print #tempfile, String(inifile(count).level, vbTab) & inifile(count).contents
            Next
        Close tempfile
        SaveFile = True
    Else
        ErrorCode = err_filenotloaded
    End If
End Function

'Saves the currently loaded data as an INI file
Public Function SaveOldFile(ByVal Filename As String) As Boolean
    On Error Resume Next
    If entrycount = 0 Then Exit Function
    Dim temp As Long, temp2 As Boolean
    temp = FreeFile
    Filename = Replace(Filename, "\\", "\")
    If Filename Like "?:\*" Then
        Open Filename For Output As temp
            temp2 = True
            If Not saveoldroot(temp) Then temp2 = False
        Close temp
    End If
    SaveOldFile = temp2
End Function

Public Function SectionAtIndex(Section As String, Index As Long)
    If entrycount = 0 Then Exit Function
    Dim temp As Long
    temp = qualifiedsectionhandle(Section)
    If temp > 0 Then
        SectionAtIndex = sectionindex(temp, Index)
    End If
End Function

Public Function SectionExists(Section As String) As Boolean
    If entrycount = 0 Then Exit Function
    Dim temp As Long
    temp = qualifiedsectionhandle(Section)
    SectionExists = False
    If temp > 0 Then SectionExists = True
End Function

Public Function Section2String(Section As String) As String
    Dim tempstr As String, temp As Long, temp2 As Long, temp3 As Long, temp4 As Long
    temp = qualifiedsectionhandle(Section)
    If temp > -1 Then
        temp2 = findroot(temp)
        temp3 = inifile(temp).level
        tempstr = inifile(temp).contents
        For temp4 = temp + 1 To temp2
            tempstr = tempstr & vbNewLine & String(inifile(temp4).level - temp3, vbTab) & inifile(temp4).contents
        Next
    End If
    Section2String = tempstr
End Function

'Sets the value of a key=value pair
Public Sub SetKey(Section As String, Key As String, Value As String, Optional Index As Long = 0)
    Dim temp As Long
    'Section = Replace(Section, "\\", "\")
    temp = qualifiedsectionhandle(Section)
    If temp = 0 Then
        CreateSection Section
        temp = qualifiedsectionhandle(Section)
    End If
    If temp > 0 Then
        If HINI_keyexists(temp, Key, Index) = True Then
            HINI_setkey temp, Key, Value, Index
        Else
            CreateKey Section, Key, Value, Index
        End If
    End If
End Sub

Public Sub SetMultiKey(Section As String, Key As String, Value As String, Optional Delimeter As String = vbNewLine)
    Dim tempstr() As String, temp As Long, temp2 As Long
    temp = qualifiedsectionhandle(Section)
    If temp > 0 Then
        tempstr = Split(Value, Delimeter)
        For temp2 = 0 To UBound(tempstr)
            If HINI_keyexists(temp, Key, temp2 + 1) = True Then
                HINI_setkey temp, Key, tempstr(temp2), temp2 + 1
            Else
                CreateKey Section, Key, tempstr(temp2), temp2 + 1
            End If
        Next
    End If
End Sub

Public Sub SetSectionContents(Section As String, Text As String, Optional IncludeTags As Boolean)
    Dim tempstr() As String, temp As Long, temp2 As Long, temp3 As Long, Currlevel As Long
    DeleteSection Section
    CreateSection Section
    temp = qualifiedsectionhandle(Section) + 1
    temp3 = temp
    Currlevel = inifile(temp).level + 1
    tempstr = Split(Text, vbNewLine)
    For temp2 = 0 To UBound(tempstr)
        If (IncludeTags And temp2 = 0) Or (IncludeTags And temp2 = UBound(tempstr)) Then temp2 = temp2 + 1
        Currlevel = HandleText(tempstr(temp2), Currlevel, temp3)
        temp3 = temp3 + 1
    Next
    For temp2 = temp To Currlevel
        insert temp3, Currlevel, "[/]"
        temp3 = temp3 + 1
    Next
End Sub

















Private Function CountRootSections() As Long
    If entrycount = 0 Then Exit Function
    Dim count As Long, sections As Long
    For count = 1 To entrycount
        If inifile(count).level = 0 Then
            If inifile(count).contents <> Empty Then
                If isnotroot(inifile(count).contents) Then
                    If isSection(inifile(count).contents) Then sections = sections + 1
                End If
            End If
        End If
    Next
    CountRootSections = sections
End Function

Private Function TextType(Text As String) As KeyType
    If isSection(Text) Then
        If isroot(Text) Then
            TextType = key_Root
        Else
            TextType = key_Root
        End If
    Else
        If isvalue(Text) Then TextType = key_KeyValuePair
    End If
End Function
Private Function isroot(Text As String) As Boolean
'If Left(text, 2) = "[/" & Right(text, 1) = "]" Then isroot = True Else isroot = False
isroot = Text = "[/]"
End Function
Private Function isSection(Value As String) As Boolean
    isSection = Left(Value, 1) = "[" And Right(Value, 1) = "]" And Len(stripsection(Value)) > 0
End Function
Private Function isnotroot(Value As String) As Boolean
isnotroot = Not Value = "[/]" 'If issection(value) = True And isroot(value) = False Then isnotroot = True
End Function
Private Function isvalue(Value As String) As Boolean
    isvalue = isSection(Value) = False And InStr(Value, "=") > 0
End Function
Private Function stripsection(Section As String) As String
    stripsection = Mid(Section, 2, Len(Section) - 2)
End Function
Private Function stripvalue(Value As String) As String
    stripvalue = Right(Value, Len(Value) - InStr(Value, "="))
End Function
Private Function stripname(Value As String) As String
    stripname = Left(Value, InStr(Value, "=") - 1)
End Function
Private Function iscomment(Value As String) As Boolean
    iscomment = Left(Value, 1) = "#" Or Left(Value, 1) = "'"
End Function
Private Sub UserControl_Initialize()
    Dim tempstr() As String
    isloaded = False
    entrycount = 0
End Sub

Private Function saveoldroot(Filenumber As Long) As Boolean
    On Error Resume Next
    If entrycount = 0 Then Exit Function
    Dim tempstr() As String, count As Long, temp As Long
    temp = countrootkeys
    If temp > 0 Then
        Print #Filenumber, "[/]"
        EnumRootKeys tempstr
        For count = 1 To temp
            Print #Filenumber, tempstr(1, count) & "=" & tempstr(2, count)
        Next
    End If
    temp = CountRootSections
    If temp > 0 Then
        EnumRootSections tempstr
        For count = 1 To temp
            saveoldsection tempstr(count), Filenumber
        Next
    End If
    saveoldroot = True
End Function

Private Sub saveoldsection(Section As String, Filenumber As Long)
    If entrycount = 0 Then Exit Sub
    Dim temp As Long, tempstr() As String, count As Long, tempstr2() As String
    temp = qualifiedsectionhandle(Section)
    If HINI_countsections(temp) > 0 Then
        HINI_EnumSections temp, tempstr
        For count = 1 To HINI_countsections(temp)
            saveoldsection Section & "\" & tempstr(count), Filenumber
        Next
    End If
    Print #Filenumber, "[" & Replace(Section, "\", "/") & "]"
    HINI_EnumKeys temp, tempstr2
    For count = 1 To HINI_countkeys(temp)
        If countinstancekeys(temp, tempstr2(1, count)) = 1 Then
            Print #Filenumber, tempstr2(1, count) & "=" & tempstr2(2, count)
        Else
            Print #Filenumber, tempstr2(1, count) & "(" & HINI_GetKeysinstance(temp, counthandlekeys(temp, count)) & ")=" & tempstr2(2, count)
        End If
    Next
    Print #Filenumber, vbNewLine 'whitespace
End Sub

Private Function HandleText(tempstr As String, Currlevel As Long, Optional Position As Long) As Long
    Dim continue As Boolean
    tempstr = Replace(tempstr, vbTab, Empty)
    If Len(tempstr) > 0 Then  '[/] sections = evil
            continue = True ' account for roots going below root level
            If isroot(tempstr) Then
                If Currlevel - 1 >= 0 Then
                    Currlevel = Currlevel - 1
                Else
                    continue = False
                End If
            End If
            If continue = True Then
                If Position = 0 Then Position = entrycount + 1 'add to end
                insert Position, Currlevel, tempstr
                If isroot(tempstr) = False And isSection(tempstr) = True Then Currlevel = Currlevel + 1
            End If
    End If
    HandleText = Currlevel
End Function

Private Function AccountForMissingRoots(ByVal Currlevel As Long) As Long
    Dim tempfile As Long, temp As Long
    If Currlevel > 0 Then 'accounts for missing roots
        temp = entrycount
        entrycount = entrycount + Currlevel
        ReDim Preserve inifile(1 To entrycount)
        For tempfile = temp + 1 To entrycount
            Currlevel = Currlevel - 1
            inifile(tempfile).level = Currlevel
            inifile(tempfile).contents = "[/]"
        Next
    End If
    AccountForMissingRoots = Currlevel
End Function

Private Function findroot(Start As Long) As Long
    If entrycount = 0 Then Exit Function
    Dim count As Long, Found As Boolean
    Found = False
    If Start > 0 Then
        If isloaded And Start > 0 Then
            For count = Start + 1 To entrycount
                If Found = False Then
                    If inifile(count).level = inifile(Start).level Then
'                If isroot(inifile(count).contents) Then
                        Found = True
                        findroot = count
                        Exit Function
'                End If
                    End If
                End If
            Next
        End If
    Else
        findroot = entrycount
    End If
End Function
Private Function HINI_countkeys(Start As Long) As Long
    If entrycount = 0 Then Exit Function
    Dim count As Long, keys As Long
    For count = Start To findroot(Start)
        If inifile(count).level = inifile(Start).level + 1 Then
            If isvalue(inifile(count).contents) Then
                keys = keys + 1
            End If
        End If
    Next
    HINI_countkeys = keys
End Function
Private Function counthandlekeys(Start As Long, Index As Long) As Long
    If entrycount = 0 Then Exit Function
    Dim count As Long, keys As Long
    For count = Start To findroot(Start)
        If inifile(count).level = inifile(Start).level + 1 Then
            If isvalue(inifile(count).contents) Then
                keys = keys + 1
                If Index = keys Then
                    counthandlekeys = count
                End If
            End If
        End If
    Next
End Function
Private Function countinstancekeys(Start As Long, Key As String) As Long
    If entrycount = 0 Then Exit Function
    Dim count As Long, keys As Long, temp As Long, temp2 As Long
    If Start = 0 Then
        temp = entrycount
        temp2 = 0
        Start = 1
    Else
        temp = findroot(Start)
        temp2 = inifile(Start).level + 1
    End If
    For count = Start To temp
        If inifile(count).level = temp2 Then
            If isvalue(inifile(count).contents) And StrComp(stripname(inifile(count).contents), LCase(Key), vbTextCompare) = 0 Then
                keys = keys + 1
            End If
        End If
    Next
    countinstancekeys = keys
End Function
Private Function HINI_GetKeysinstance(Start As Long, handle As Long) As Long
    If entrycount = 0 Then Exit Function
    Dim count As Long, keys As Long, temp As String
    temp = LCase(stripname(inifile(handle).contents))
    For count = Start To handle
        If inifile(count).level = inifile(Start).level + 1 Then
            If isvalue(inifile(count).contents) And StrComp(stripname(inifile(count).contents), temp, vbTextCompare) = 0 Then
                keys = keys + 1
            End If
        End If
    Next
    HINI_GetKeysinstance = keys
End Function
Private Function HINI_countsections(Start As Long) As Long
    If entrycount = 0 Then Exit Function
    Dim count As Long, sections As Long
    If Start > 0 Then
        For count = Start To findroot(Start)
            If inifile(count).level = inifile(Start).level + 1 Then
                If isnotroot(inifile(count).contents) Then
                    If isSection(inifile(count).contents) Then sections = sections + 1
                End If
            End If
        Next
    Else
        sections = CountRootSections
    End If
    HINI_countsections = sections
End Function
Private Function MiniLike(filter As String, expression As String) As Boolean
    Dim temp As Boolean, Ls As String, Rs As String, Lb As Boolean, Rb As Boolean
    temp = expression Like filter
    If Not temp Then
        Ls = Left(filter, 1)
        Rs = Right(filter, 1)
        Lb = Ls = "*" Or Ls = "?"
        Rb = Rs = "*" Or Rs = "?"
        If Lb And Rb Then
            temp = expression Like Mid(filter, 2, Len(filter) - 2)
        Else
            If Lb Then temp = expression Like Right(filter, Len(filter) - 1)
            If Rb Then temp = expression Like Left(filter, Len(filter) - 1)
        End If
    End If
    MiniLike = temp
End Function




Private Function EnumRootSections(strArray, Optional Key As String, Optional Index As Long) As Long
    If entrycount = 0 Then Exit Function
    Dim count As Long, sections As Long
    For count = 1 To entrycount
        If inifile(count).level = 0 Then
            If Len(inifile(count).contents) > 0 Then
                If isnotroot(inifile(count).contents) Then
                    If isSection(inifile(count).contents) Then
                        sections = sections + 1
                        If Len(Key) = 0 Then
                            ReDim Preserve strArray(1 To sections)
                            strArray(sections) = stripsection(inifile(count).contents)
                        Else
                            ReDim Preserve strArray(1 To 2, 1 To sections)
                            strArray(1, sections) = stripsection(inifile(count).contents)
                            strArray(2, sections) = HINI_GetKey(count, Key, Index)
                        End If
                    End If
                End If
            End If
        End If
    Next
    EnumRootSections = sections
End Function
Private Function handlerootsections(Section As String) As Long
    If entrycount = 0 Then Exit Function
    Dim count As Long
    For count = 1 To entrycount
        If inifile(count).level = 0 Then
            If isnotroot(inifile(count).contents) And isSection(inifile(count).contents) Then
                If StrComp(stripsection(inifile(count).contents), Section, vbTextCompare) = 0 Then
                    handlerootsections = count
                    Exit Function
                End If
            End If
        End If
    Next
End Function

Private Function HINI_EnumSections(Start As Long, strArray, Optional Key As String, Optional Index As Long) As Long
    If entrycount = 0 Then Exit Function
    Dim count As Long, sections As Long
    For count = Start To findroot(Start)
        If inifile(count).level = inifile(Start).level + 1 Then
            If isnotroot(inifile(count).contents) And isSection(inifile(count).contents) Then
                sections = sections + 1
                If Len(Key) = 0 Then
                    ReDim Preserve strArray(1 To sections)
                    strArray(sections) = stripsection(inifile(count).contents)
                Else
                    ReDim Preserve strArray(1 To 2, 1 To sections)
                    strArray(1, sections) = stripsection(inifile(count).contents)
                    strArray(2, sections) = HINI_GetKey(count, Key, Index)
                    
                End If
            End If
        End If
    Next
    HINI_EnumSections = sections
End Function

Private Function EnumRootKeys(strArray) As Long
    If entrycount = 0 Then Exit Function
    Dim count As Long, keys As Long
    For count = 1 To entrycount
        If inifile(count).level = 0 Then
            If isvalue(inifile(count).contents) Then
                keys = keys + 1
                ReDim Preserve strArray(1 To 2, 1 To keys) As String
                strArray(1, keys) = stripname(inifile(count).contents)
                strArray(2, keys) = stripvalue(inifile(count).contents)
            End If
        End If
    Next
    EnumRootKeys = keys
End Function
Private Function countrootkeys() As Long
    If entrycount = 0 Then Exit Function
    Dim count As Long, keys As Long
    For count = 1 To entrycount
        If inifile(count).level = 0 Then
            If isvalue(inifile(count).contents) Then
                keys = keys + 1
            End If
        End If
    Next
    countrootkeys = keys
End Function
Private Function HINI_EnumKeys(Start As Long, strArray, Optional Recursive As Boolean) As Long
    If entrycount = 0 Then Exit Function
    Dim count As Long, keys As Long, CurrSection As String
    If Start > 0 Then
        For count = Start + 1 To findroot(Start)
            If Recursive Then
                If isvalue(inifile(count).contents) Then
                        keys = keys + 1
                        ReDim Preserve strArray(0 To 2, 1 To keys) As String
                        strArray(0, keys) = CurrSection
                        strArray(1, keys) = stripname(inifile(count).contents)
                        strArray(2, keys) = stripvalue(inifile(count).contents)
                        
                ElseIf isroot(inifile(count).contents) Then
                    If InStr(CurrSection, "\") = 0 Then
                        CurrSection = Empty
                    Else
                        CurrSection = Left(CurrSection, InStrRev(CurrSection, "\") - 1)
                    End If
                        
                ElseIf isSection(inifile(count).contents) Then
                    If Len(CurrSection) = 0 Then
                        CurrSection = stripsection(inifile(count).contents)
                    Else
                        CurrSection = CurrSection & "\" & stripsection(inifile(count).contents)
                    End If
                End If
            Else
                If inifile(count).level = inifile(Start).level + 1 Then
                    If isvalue(inifile(count).contents) Then
                        keys = keys + 1
                        ReDim Preserve strArray(1 To 2, 1 To keys) As String
                        strArray(1, keys) = stripname(inifile(count).contents)
                        strArray(2, keys) = stripvalue(inifile(count).contents)
                    End If
                End If
            End If
        Next
        HINI_EnumKeys = keys
    Else
        HINI_EnumKeys = EnumRootKeys(strArray)
    End If
End Function
Private Function HINI_sectionexists(Start As Long, Section As String) As Boolean
    If entrycount = 0 Then Exit Function
    Dim tempstr() As String, count As Long, temp As Long
    If Start > 0 Then
        temp = HINI_EnumSections(Start, tempstr)
        'temp = HINI_countsections(start)
    Else
        temp = EnumRootSections(tempstr)
        'temp = CountRootSections
    End If
    HINI_sectionexists = False
    For count = 1 To temp
        If StrComp(tempstr(count), Section, vbTextCompare) = 0 Then
            HINI_sectionexists = True
            Exit For
        End If
    Next
End Function

Private Function sectionindex(Start As Long, Index As Long) As String
    If entrycount = 0 Or Index = 0 Then Exit Function
    Dim tempstr() As String, temp As Long
    If Start > 0 Then
        HINI_EnumSections Start, tempstr
        temp = HINI_countsections(Start)
    Else
        EnumRootSections tempstr
        temp = CountRootSections
    End If
    If Index <= temp And temp > 0 Then
        sectionindex = tempstr(Index)
    End If
End Function

Private Function HINI_keyexists(Start As Long, Key As String, Optional Index As Long = 0) As Boolean
    If entrycount = 0 Then Exit Function
    Dim tempstr() As String, count As Long, count2 As Long
    If Start > 0 Then HINI_EnumKeys Start, tempstr Else EnumRootKeys tempstr
    HINI_keyexists = False
    For count = 1 To HINI_countkeys(Start)
        If StrComp(tempstr(1, count), Key, vbTextCompare) = 0 Then
            count2 = count2 + 1
            If count2 = Index Or Index = 0 Then
                HINI_keyexists = True
            End If
        End If
    Next
End Function

Private Function keyindex(Start As Long, Index As Long, Optional Side As Long = 2) As String
    If entrycount = 0 Then Exit Function
    Dim tempstr() As String
    HINI_EnumKeys Start, tempstr
    If Index > 0 And Index <= HINI_countkeys(Start) Then
        keyindex = tempstr(Side, Index)
    End If
End Function

Private Function keyhandle(Start As Long, Key As String, Optional Index As Long = 0) As Long
    If entrycount = 0 Then Exit Function
    Dim count As Long, count2 As Long, Found As Boolean
    Found = False
    For count = Start To findroot(Start)
        If inifile(count).level = inifile(Start).level + 1 And Found = False Then
            If isvalue(inifile(count).contents) Then
                If StrComp(stripname(inifile(count).contents), Key, vbTextCompare) = 0 Then
                    count2 = count2 + 1
                    If count2 = Index Or Index = 0 Then
                        keyhandle = count
                        Found = True
                    End If
                End If
            End If
        End If
    Next
End Function
Private Function newrootsection(Section As String)
    If HINI_sectionexists(0, Section) = False Then
        isloaded = True
        entrycount = entrycount + 2
        If entrycount > 2 Then ReDim Preserve inifile(1 To entrycount)
        If entrycount = 2 Then ReDim inifile(1 To entrycount)
        inifile(entrycount - 1).level = 0
        inifile(entrycount - 1).contents = "[" & Section & "]"
        inifile(entrycount).level = 0
        inifile(entrycount).contents = "[/]"
    End If
End Function

Private Function HINI_countmultikey(Start As Long, Key As String) As Long
If entrycount = 0 Then Exit Function
    On Error Resume Next
Dim count As Long, temp As Long, count2 As Long
If Start = 0 Then
    temp = 0
    Start = 1
Else
    temp = inifile(Start).level + 1
End If
For count = Start To findroot(Start)
    If inifile(count).level = temp Then
        If isvalue(inifile(count).contents) Then
            If StrComp(stripname(inifile(count).contents), Key, vbTextCompare) = 0 Then
                count2 = count2 + 1
            End If
        End If
    End If
Next
HINI_countmultikey = count2
End Function
Private Function HINI_GetKey(Start As Long, Key As String, Optional Index As Long = 0) As String
If entrycount = 0 Then Exit Function
    Dim count As Long, count2 As Long, temp As Long, Found As Boolean
    'This code is by Techni Myoko, and anyone who claims otherwise is a liar
Found = False
If Start = 0 Then
    temp = 0
    Start = 1
Else
    temp = inifile(Start).level + 1
End If
For count = Start To findroot(Start)
    If inifile(count).level = temp And Found = False Then
        If isvalue(inifile(count).contents) Then
            If StrComp(stripname(inifile(count).contents), Key, vbTextCompare) = 0 Then
                count2 = count2 + 1
                If Index = 0 Or count2 = Index Then
                    HINI_GetKey = stripvalue(inifile(count).contents)
                    Found = True
                End If
            End If
        End If
    End If
Next
End Function

Private Function ChkSection(ByVal Path As String, ByVal File As String) As String
    If Left(File, 1) = "\" Then File = Right(File, Len(File) - 1)
    If Len(Path) = 0 Then
        ChkSection = File
    Else
        If Right(Path, 1) = "\" Then Path = Left(Path, Len(Path) - 1)
        ChkSection = Path & "\" & File
    End If
End Function


Private Sub HINI_setkey(Start As Long, Key As String, Value As String, Optional Index As Long = 0)
Dim count As Long, count2 As Long, Found As Boolean
Found = False
If HINI_keyexists(Start, Key) = True Then
For count = Start To findroot(Start)
    If inifile(count).level = inifile(Start).level + 1 And Found = False Then
        If isvalue(inifile(count).contents) Then
            If StrComp(stripname(inifile(count).contents), Key, vbTextCompare) = 0 Then
                count2 = count2 + 1
                If Index = 0 Or count2 = Index Then
                    inifile(count).contents = stripname(inifile(count).contents) & "=" & Value
                    Found = True
                End If
            End If
        End If
    End If
Next
End If
End Sub
Private Function sectionhandle(Start As Long, Section As String) As Long
If entrycount = 0 Then Exit Function
    Dim count As Long
If Start > 0 Then
For count = Start To findroot(Start)
    If inifile(count).level = inifile(Start).level + 1 Then
        If isSection(inifile(count).contents) Then
            If StrComp(stripsection(inifile(count).contents), Section, vbTextCompare) = 0 Then
                sectionhandle = count
            End If
        End If
    End If
Next
Else
    sectionhandle = handlerootsections(Section)
End If
End Function

Private Function qualifiedsectionhandle(Section As String) As Long
If entrycount = 0 Then Exit Function
Dim tempstr() As String, count As Long, temp As Long, exists As Boolean
If Len(Section) > 0 Then
    tempstr = Split(Section, "\")
    exists = HINI_sectionexists(0, tempstr(0))
    If exists Then
        temp = handlerootsections(tempstr(0))
        For count = 1 To UBound(tempstr)
            exists = exists And HINI_sectionexists(temp, tempstr(count))
            If exists = True Then
                temp = sectionhandle(temp, tempstr(count))
            End If
        Next
    End If
    If exists = True Then qualifiedsectionhandle = temp
End If
End Function
Private Sub HINI_RenameKey(Start As Long, Key As String, name As String, Optional Index As Long = 0, Optional newindex As Long = 0)
If entrycount = 0 Then Exit Sub
    Dim temp As Long
temp = keyhandle(Start, Key, Index)
If temp > 0 And HINI_keyexists(Start, name, newindex) = False Then
    inifile(temp).contents = name & "=" & stripvalue(inifile(temp).contents)
End If
End Sub

Private Sub removesection(Start As Long, Optional LeaveTags As Boolean)
    If entrycount > 0 Then
        If LeaveTags Then
            removerange Start + 1, findroot(Start) - 1
        Else
            removerange Start, findroot(Start)
        End If
    End If
End Sub
Private Sub HINI_renamesection(Start As Long, Section As String)
If entrycount = 0 Then Exit Sub
        If HINI_sectionexists(Start, Section) = False And Section <> "/" Then
        inifile(Start).contents = "[" & Replace(Section, "\", "/") & "]"
    End If
End Sub
Private Function concatenate(strArray, Delimeter As String) As String
'Dim temp As String, count As Long
'For count = LBound(STRarray) To UBound(STRarray)
'    If temp <> Empty Then temp = temp & delimeter
'    temp = temp & STRarray(count)
'Next
concatenate = Join(strArray, Delimeter)
End Function

Private Sub HINI_createsection(ByVal Start As Long, Section As String)
    Dim Finish As Long
    Section = Replace(Section, "\", "/")
    If Start > 0 Then
    Finish = findroot(Start)
    If HINI_sectionexists(Start, Section) = False And Section <> "/" And Section <> Empty Then
        insert Finish, inifile(Start).level + 1, "[/]"
        insert Finish, inifile(Start).level + 1, "[" & Section & "]"
    End If
    Else
        newrootsection Section
    End If
End Sub
Private Sub insert(Start As Long, level As Long, contents As String)
    Dim count As Long
    'Redimension ini item array
    entrycount = entrycount + 1
    ReDim Preserve inifile(1 To entrycount)
    'Move all after the insertion point up one
    For count = entrycount - 1 To Start Step -1
        inifile(count + 1).contents = inifile(count).contents
        inifile(count + 1).level = inifile(count).level
    Next
    'Insert
    inifile(Start).contents = contents
    inifile(Start).level = level
End Sub
Private Sub removerange(Top As Long, Bottom As Long)
    If entrycount = 0 Then Exit Sub
    Dim range As Long, count As Long
    If Top > 0 And Bottom > 0 And Top <= entrycount And Bottom <= entrycount Then 'And Top >= Bottom Then
        range = Bottom + 1 - Top
        If range > 0 Then
        entrycount = entrycount - range
        For count = Top To entrycount
            inifile(count).contents = inifile(count + range).contents
            inifile(count).level = inifile(count + range).level
        Next
        If entrycount > 0 Then
            ReDim Preserve inifile(1 To entrycount)
        Else
            ReDim inifile(entrycount)
        End If
        End If
    End If
End Sub

Private Sub ReDimPreserve(ByRef count As Long, strArray() As entry, level As Long, contents As String)
    count = count + 1
    If count = 1 Then
        ReDim strArray(1 To 1) As entry
    Else
        ReDim Preserve strArray(1 To count) As entry
    End If
    strArray(count).level = level
    strArray(count).contents = LTrim(contents)
End Sub




Private Sub XML_ParseTag(Text As String, ByRef Currlevel As Long)
    Dim temp As Long, count As Long, tempstr() As String, tempstr2 As String
    If Left(Text, 2) = "</" Then
        Currlevel = Currlevel - 1
        ReDimPreserve entrycount, inifile, Currlevel, "[/]"
    Else
        count = XML_SplitTag(Text, tempstr)
        ReDimPreserve entrycount, inifile, Currlevel, "[" & tempstr(0) & "]"
        Currlevel = Currlevel + 1
        
        For temp = 1 To count - 1
            tempstr2 = XML_Trim(XML_TagValue(tempstr(temp)))
            If Len(tempstr2) > 0 Then ReDimPreserve entrycount, inifile, Currlevel, XML_TagName(tempstr(temp)) & "=" & tempstr2
        Next
        
        If XML_SelfTerminating(Text) Then
            Currlevel = Currlevel - 1
            ReDimPreserve entrycount, inifile, Currlevel, "[/]"
        End If
    End If
End Sub
Private Function XML_TagName(ByVal Text As String) As String
    Dim temp As Long
    temp = InStr(Text, "=")
    If temp > 0 Then Text = Left(Text, temp - 1)
    XML_TagName = XML_Trim(Text)
End Function
Private Function XML_TagValue(ByVal Text As String) As String
    Dim temp As Long
    temp = InStr(Text, "=")
    If temp > 0 Then
        Text = XML_Trim(Right(Text, Len(Text) - temp))
        If Left(Text, 1) = """" And Right(Text, 1) = """" Then Text = Mid(Text, 2, Len(Text) - 2)
        XML_TagValue = Text
    End If
End Function
Private Function XML_SplitTag(ByVal Text As String, temparr) As Long
    Dim temp As Long, count As Long, tempstr As String
    Text = XML_Trim(Text)
    Text = XML_Trim(Mid(Text, 2, Len(Text) - 2))
    If Left(Text, 1) = "?" And Right(Text, 1) = "?" Then Text = Mid(Text, 2, Len(Text) - 2)
    If Left(Text, 1) = "/" Then Text = Right(Text, Len(Text) - 1)
    If Right(Text, 1) = "/" Then Text = Left(Text, Len(Text) - 1)
    Text = XML_Trim(Text)
     
    Do Until Len(Text) = 0
        tempstr = XML_Grabtagword(Text)
        If Len(tempstr) = 0 Then
            Text = Empty
        Else
            count = count + 1
            ReDim Preserve temparr(count)
            temparr(count - 1) = tempstr
            Text = Right(Text, Len(Text) - Len(tempstr))
        End If
        Text = XML_Trim(Text)
    Loop
    XML_SplitTag = count
End Function
Private Function XML_Grabtagword(ByVal Text As String) As String
    Dim temp As Long, isinquotes As Boolean, endit As Boolean
    Text = XML_Trim(Text)
    For temp = 1 To Len(Text)
        If isinquotes Then
            If Mid(Text, temp, 1) = """" Then
                isinquotes = False
                endit = True
            End If
        Else
            Select Case Mid(Text, temp, 1)
                Case """": isinquotes = True
                Case " ", ">": endit = True
            End Select
        End If
        
        If endit Then
            XML_Grabtagword = XML_Trim(Left(Text, temp))
            Exit Function
        End If
    Next
    If Not endit Then XML_Grabtagword = Text
End Function
Private Function XML_Grabword(Text As String) As String
    Dim temp As Long, temp2 As Long
    temp = InStr(Text, "<")
    If temp > 0 Then
        If temp = 1 Then
            temp2 = InStr(Text, ">")
            XML_Grabword = Left(Text, temp2)
        Else
            XML_Grabword = Left(Text, temp - 1)
        End If
    End If
End Function
Private Function XML_Trim(ByVal Text As String) As String
    Dim temp As Long
    For temp = 1 To Len(Text)
        Select Case Left(Text, 1)
            Case " ", vbNewLine, Chr(9), Chr(10): Text = Right(Text, Len(Text) - 1)
            Case Else: Exit For
        End Select
    Next
    For temp = Len(Text) To 1 Step -1
        Select Case Right(Text, 1)
            Case " ", vbNewLine, Chr(9), Chr(10): Text = Left(Text, Len(Text) - 1)
            Case Else: Exit For
        End Select
    Next
    XML_Trim = Trim(Text)
End Function
Private Function XML_SelfTerminating(Text As String) As Boolean
    XML_SelfTerminating = Left(Text, 2) = "<?" And Right(Text, 2) = "?>" Or Right(Text, 2) = "/>"
End Function
Private Sub XML_Compress()
    If entrycount = 0 Then Exit Sub
    Dim count As Long
    For count = entrycount - 2 To 1 Step -1
        If isSection(inifile(count).contents) Then
            If isvalue(inifile(count + 1).contents) Then
                If StrComp(stripname(inifile(count + 1).contents), "Node", vbTextCompare) = 0 Then
                    If isroot(inifile(count + 2).contents) Then
                        inifile(count).contents = stripsection(inifile(count).contents) & "=" & stripvalue(inifile(count + 1).contents)
                        removerange count + 1, count + 2
                    End If
                End If
            End If
        End If
    Next
End Sub


Private Sub removedoublekey(Start As Long)
If entrycount = 0 Then Exit Sub
        Dim count As Long, count2 As Long, count3 As Long
    Dim root As Long
    root = findroot(Start) - 1
    For count = Start + 1 To root
        If inifile(count).level = inifile(Start).level + 1 Then
            If isvalue(inifile(count).contents) Then
                count3 = 0
                For count2 = count + 1 To root
                    If inifile(count2).level = inifile(count).level Then
                        If isvalue(inifile(count2).contents) Then
                            If StrComp(stripname(inifile(count2).contents), stripname(inifile(count).contents), vbTextCompare) = 0 Then
                                count3 = count3 + 1
                                inifile(count2).contents = Left(inifile(count2).contents, InStr(inifile(count2).contents, "=") - 1) & "(" & count3 & ")=" & Right(inifile(count2).contents, Len(inifile(count2).contents) = InStr(inifile(count2).contents, "="))
                            End If
                        End If
                    End If
                Next
                If count3 > 0 Then
                    inifile(count).contents = Left(inifile(count).contents, InStr(inifile(count).contents, "=") - 1) & "(0)=" & Right(inifile(count).contents, Len(inifile(count).contents) = InStr(inifile(count).contents, "="))
                End If
            End If
        End If
    Next
End Sub
Private Sub removedoublessection(Start As Long)
If entrycount = 0 Then Exit Sub
        Dim count As Long, count2 As Long, count3 As Long
    Dim root As Long
    root = findroot(Start) - 1
   ' MsgBox root
    For count = Start + 1 To root
        If inifile(count).level = inifile(Start).level + 1 Then
            If isnotroot(inifile(count).contents) Then
                count3 = 0
                For count2 = count + 1 To root
                    If inifile(count2).level = inifile(count).level Then
                        If isnotroot(inifile(count2).contents) Then
                            If StrComp(stripsection(inifile(count2).contents), stripsection(inifile(count).contents), vbTextCompare) = 0 Then
                                count3 = count3 + 1
                                inifile(count2).contents = Left(inifile(count2).contents, Len(inifile(count2).contents) - 1) & "(" & count3 & ")]"
                            End If
                        End If
                    End If
                Next
                If count3 > 0 Then
                    inifile(count).contents = Left(inifile(count).contents, Len(inifile(count).contents) - 1) & "(0)]"
                End If
            End If
        End If
    Next
End Sub
Private Sub UserControl_Resize()
UserControl.Width = 480
UserControl.Height = UserControl.Width
End Sub
Public Property Let Tag(Text As String)
    stag = Text
End Property
Public Property Get Tag() As String
    Tag = stag
End Property
Private Sub CreateKey(Section As String, Key As String, Value As String, Optional Index As Long = 0)
    Dim temp As Long
    temp = qualifiedsectionhandle(Section)
    If temp > 0 Then
        temp = findroot(temp)
        If HINI_keyexists(temp, Key, Index) = False Then
            insert temp, inifile(temp).level + 1, Key & "=" & Value
        Else
            SetKey Section, Key, Value, Index
        End If
    End If
End Sub
