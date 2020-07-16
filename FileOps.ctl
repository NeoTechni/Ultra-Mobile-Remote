VERSION 5.00
Begin VB.UserControl FileOps 
   ClientHeight    =   525
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   840
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   525
   ScaleWidth      =   840
   Begin VB.Timer TimerDelete 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   0
      Top             =   0
   End
   Begin VB.FileListBox Filmain 
      Height          =   285
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   255
   End
   Begin VB.DirListBox DirMain 
      Height          =   315
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   255
   End
End
Attribute VB_Name = "FileOps"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Dim m_FromFile As String, m_ToFile As String, isInProgress As Boolean, LastUpdate As Long, CUpdate As Long
Dim cOp As Long, newDest As String, doCancel As Boolean

Const DefaultBuffer As Long = 1024 * 16 'kilobytes
Const MaxLong As Long = 2147483647

Const SHERB_NOCONFIRMATION = &H1
Const SHERB_NOPROGRESSUI = &H2
Const SHERB_NOSOUND = &H4

Const FO_COPY = &H2
Const FO_DELETE = &H3
Const FO_MOVE = &H1
Const FO_RENAME = &H4
Const FOF_ALLOWUNDO = &H40 'By adding the FOF_ALLOWUNDO flag you can move a file to the Recycle Bin instead of deleting it.
Const FOF_SILENT = &H4
Const FOF_NOCONFIRMATION = &H10
Const FOF_RENAMEONCOLLISION = &H8
Const FOF_NOCONFIRMMKDIR = &H200
Const FOF_NOERRORUI As Long = &H400
Const FOF_FILESONLY = &H80

Private Type ULARGE_INTEGER
    LowPart As Long
    HighPart As Long
End Type

Private Type SHQUERYRBINFO
    cbSize As Long
    i64Size As ULARGE_INTEGER
    i64NumItems As ULARGE_INTEGER
End Type

Private Type SHFILEOPSTRUCT
    hwnd      As Long
    wFunc     As Long
    pFrom     As String
    pTo       As String
    fFlags    As Integer
    fAborted  As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function SHEmptyRecycleBin Lib "shell32.dll" Alias "SHEmptyRecycleBinA" (ByVal hwnd As Long, ByVal pszRootPath As String, ByVal dwFlags As Long) As Long
Private Declare Function SHUpdateRecycleBinIcon Lib "shell32.dll" () As Long
Private Declare Function SHQueryRecycleBin Lib "shell32.dll" Alias "SHQueryRecycleBinA" (ByVal pszRootPath As String, pSHQueryRBInfo As SHQUERYRBINFO) As Long
      
Private Enum ShellExecuteConstants
    SW_HIDE
    SW_SHOWNORMAL
    SW_SHOWMINIMIZED
    SW_SHOWMAXIMIZED
    SW_SHOWNOACTIVATE
    SW_SHOW
    SW_MINIMIZE
    SW_SHOWMINNOACTIVE
    SW_SHOWNA
    SW_RESTORE
    SW_SHOWDEFAULT
End Enum

Dim FileQueue() As String, FileCount As Long, DirQueue() As String, DirCount As Long
Dim OriginalDir As String, FileSize As Long, DidSomething As Boolean

Public Event CopyProgressReport(FilesCopied As Long, EstimatedSecondsRemaining As Long)
Public Event DeleteProgressReport(FilesDeleted As Long, EstimatedSecondsRemaining As Long)

Public Event FileProgressReport(BytesRead As Long, BytesTotal As Long, EstimatedSecondsRemaining As Long)
Public Event FileComplete()
Public Event FileStarted(Filename As String)

Public Event DeleteComplete()
Public Event CopyComplete()

Public Cut As Boolean

Public Sub Search(SearchQuery As String)
    Dim Temp As Long, File As String
    For Temp = FileCount - 1 To 0 Step -1
        File = FileQueue(Temp)
        File = Right(File, Len(File) - InStrRev(File, "\"))
        If Not SearchText(File, SearchQuery) Then Unqueue Temp
    Next
End Sub

Public Function QueuedItem(Index As Long) As String
    If Index > -1 And Index < FileCount Then QueuedItem = FileQueue(Index)
End Function

Private Sub Unqueue(Index As Long)
    Dim Temp As Long
    For Temp = Index + 1 To FileCount - 1
        FileQueue(Temp - 1) = FileQueue(Temp)
    Next
    FileCount = FileCount - 1
    If FileCount = 0 Then
        ReDim FileQueue(0)
    Else
        ReDim Preserve FileQueue(FileCount)
    End If
End Sub

Public Function UnqueueFileRecursive(Filename As String) As Long
    Dim Temp As Long, length As Long, count As Long
    length = Len(Filename) 'allow for recursive delete
    For Temp = FileCount - 1 To 0 Step -1
        If StrComp(Filename, Left(FileQueue(Temp), length), vbTextCompare) = 0 Then
            RemoveItem FileQueue, FileCount, Temp
            count = count + 1
        End If
    Next
    For Temp = DirCount - 1 To 0 Step -1
        If StrComp(Filename, Left(DirQueue(Temp), length), vbTextCompare) = 0 Then
            RemoveItem DirQueue, DirCount, Temp
            count = count + 1
        End If
    Next
    UnqueueFileRecursive = count
End Function

Public Function UnqueueFile(Filename As String)
    Dim Temp As Long
    Temp = FindFile(Filename)
    If Temp > -1 Then Unqueue Temp
End Function

Private Function FindFile(Filename As String) As Long
    Dim Temp As Long
    FindFile = -1
    For Temp = 0 To FileCount - 1
        If StrComp(Filename, FileQueue(Temp), vbTextCompare) = 0 Then
            FindFile = Temp
            Exit For
        End If
    Next
End Function


Private Function SearchText(Text As String, SearchQuery As String) As Boolean
    If IsAPattern(SearchQuery) Then
        SearchText = IsLike(Text, SearchQuery)
    Else
        SearchText = InStr(1, Text, SearchQuery, vbTextCompare) = 0
    End If
End Function

Private Function IsAPattern(Text As String) As Boolean
    IsAPattern = InStr(Text, "?") > 0 Or InStr(Text, "*") > 0
End Function
Private Function IsLike(Text As String, Expression As String) As Boolean 'islike("*.exe", "test.exe")
    Dim tempstr() As String, count As Long
    Expression = LCase(Expression)
    Text = LCase(Text)
    If InStr(Expression, ";") > 0 Then
        tempstr = Split(Expression, ";")
        For count = 0 To UBound(tempstr)
            If Text Like tempstr(count) Then
                IsLike = True
                Exit For
            End If
        Next
    Else
        IsLike = Text Like Expression
    End If
End Function

Public Function CanUndo() As Boolean
    If DidSomething Then
        CanUndo = cOp <> FO_DELETE 'FIX WHEN UNDELETE IS POSSIBLE!
    End If
End Function


Private Sub TimerDelete_Timer()
    Dim ETA As Long, Delta As Long, Temp As Long
    Temp = CUpdate
    Delta = LastUpdate - CUpdate
    If Delta = 0 Then
        ETA = MaxLong
    Else
        ETA = Temp / Delta
    End If
    Select Case cOp
        Case FO_DELETE
            RaiseEvent DeleteProgressReport(Delta, ETA)
        Case FO_MOVE, FO_COPY
            RaiseEvent CopyProgressReport(Delta, ETA)
    End Select
    LastUpdate = Temp
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    Name AsyncProp.Value As m_ToFile
    Cancel
    RaiseEvent FileComplete
End Sub
Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
    RaiseEvent FileProgressReport(AsyncProp.BytesRead, AsyncProp.BytesMax, 0)
    CUpdate = CUpdate - AsyncProp.BytesRead
End Sub

Public Sub ClearQueue()
    FileCount = 0
    ReDim FileQueue(0)
    DirCount = 0
    ReDim DirQueue(0)
    OriginalDir = Empty
    FileSize = 0
    newDest = Empty
    DidSomething = False
    
    Dirmain.Refresh
    Filmain.Refresh
End Sub



Private Sub RemoveItem(List, count As Long, Item As Long)
    Dim Temp As Long
    For Temp = Item To count - 2
        List(Temp) = List(Temp + 1)
    Next
    count = count - 1
    If count = 0 Then ReDim List(0) Else ReDim Preserve List(count)
End Sub

Public Function QueuedItems() As Long
    QueuedItems = FileCount + DirCount
End Function

Public Function QueueFile(Filename As String, Optional Recursive As Boolean) As Boolean
    Dim continue As Boolean, isdir As Boolean
    If Len(OriginalDir) = 0 Then
        continue = True
        'If IsADir(Filename) Then
            'OriginalDir = Filename
        'Else
            OriginalDir = Left(Filename, InStrRev(Filename, "\") - 1)
        'End If
    Else
        continue = StrComp(Left(Filename, Len(OriginalDir)), OriginalDir, vbTextCompare) = 0
    End If
    If continue Then
        isdir = IsADir(Filename)
        AddFile Filename, isdir, Recursive
        QueueFile = True
    End If
End Function

Public Sub Cancel()
    ' If no download is underway, complain.
    If isInProgress Then
        ' Cancel the copy.
        CancelAsyncRead m_FromFile
        isInProgress = False
        ' We are no longer copying.
        m_FromFile = Empty
        m_ToFile = Empty
    End If
    doCancel = True
End Sub


'operations: Cut/move, copy, delete, open, undo
Public Sub Delete()
    Dim Temp As Long, StartTime As Date, Files As Long
    StartTime = Now
    doCancel = False
    LastUpdate = QueuedItems
    cOp = FO_DELETE
    'TimerDelete.Enabled = True
    CUpdate = QueuedItems
    For Temp = 0 To FileCount - 1 'To 0 Step -1
        DeleteFile FileQueue(Temp), True
        CUpdate = CUpdate - 1
        
        Files = Files + 1
        DoProgress StartTime, Files, True
        
        If doCancel Then Exit Sub
        DoEvents
    Next
    For Temp = 0 To DirCount - 1 'To 0 Step -1
        DeleteFile DirQueue(Temp), True
        CUpdate = CUpdate - 1
        
        Files = Files + 1
        DoProgress StartTime, Files, True
        
        If doCancel Then Exit Sub
        DoEvents
    Next
    TimerDelete.Enabled = False
    DidSomething = True
    RaiseEvent DeleteComplete
End Sub

Private Sub DoProgress(StartTime As Date, CurrentFile As Long, Optional Delete As Boolean)
    Dim ETA As Long, Diff As Long, FilesPerSecond As Single, FilesRemaining As Long
    
    Diff = DateDiff("s", StartTime, Now)
    If Diff > 0 Then
        FilesPerSecond = CurrentFile / Diff
        FilesRemaining = QueuedItems - CurrentFile
        If FilesPerSecond > 0 Then
            ETA = FilesRemaining / FilesPerSecond
            If Delete Then
                RaiseEvent DeleteProgressReport(CurrentFile, ETA)
            Else
                RaiseEvent CopyProgressReport(CurrentFile, ETA)
            End If
        End If
    End If
End Sub

Public Sub CopyTo(Destination As String, Optional Cut As Boolean, Optional DoRAW As Boolean = True)
    Dim Temp As Long, Dest As String
    If Cut Then cOp = FO_MOVE Else cOp = FO_COPY
    newDest = Destination
    LastUpdate = FileSize
    CUpdate = FileSize
    doCancel = False
    TimerDelete.Enabled = True
    CUpdate = QueuedItems
    For Temp = 0 To FileCount - 1 'To 0 Step -1
        Dest = Destination & Right(FileQueue(Temp), Len(FileQueue(Temp)) - Len(OriginalDir))
        MoveFile FileQueue(Temp), Dest, Cut
        CUpdate = CUpdate - 1
        
        If doCancel Then Exit Sub
        DoEvents
    Next
    
    For Temp = 0 To DirCount - 1 'To 0 Step -1
        Dest = Destination & Right(DirQueue(Temp), Len(DirQueue(Temp)) - Len(OriginalDir))
        MkPath Dest
        If Cut Then DeleteFile DirQueue(Temp), False
        CUpdate = CUpdate - 1
        
        If doCancel Then Exit Sub
        DoEvents
    Next
    TimerDelete.Enabled = False
    DidSomething = True
    
    RaiseEvent CopyComplete
End Sub

Public Sub OpenFiles(Optional Verb As String = "Open")
    Dim Temp As Long
    doCancel = False
    For Temp = 0 To FileCount - 1 ' To 0 Step -1
        ShellFile UserControl.hwnd, FileQueue(Temp), Verb
        If doCancel Then Exit Sub
        DoEvents
    Next
    For Temp = 0 To DirCount - 1
        ShellFile UserControl.hwnd, DirQueue(Temp), Verb
        If doCancel Then Exit Sub
        DoEvents
    Next
End Sub

Public Sub Rename(Source As String, Destination As String)
    ClearQueue
    QueueFile Source, False
    CopyTo Destination, True
    DidSomething = True
End Sub

Public Function Undo() As Boolean
    Dim Temp As Long, Source As String, tempstr As String
    If DidSomething Then
    doCancel = False
    Select Case cOp
        Case FO_DELETE
            'Undelete
            
        Case FO_MOVE
            tempstr = OriginalDir
            For Temp = 0 To FileCount - 1
                Source = newDest & Right(FileQueue(Temp), Len(FileQueue(Temp)) - Len(OriginalDir))
                FileQueue(Temp) = Source
                If doCancel Then Exit Function
                DoEvents
            Next
            OriginalDir = newDest
            CopyTo tempstr, True
        Case FO_COPY
            For Temp = 0 To FileCount - 1
                Source = newDest & Right(FileQueue(Temp), Len(FileQueue(Temp)) - Len(OriginalDir))
                FileQueue(Temp) = Source
            Next
            For Temp = 0 To DirCount - 1
                Source = newDest & Right(DirQueue(Temp), Len(DirQueue(Temp)) - Len(OriginalDir))
                DirQueue(Temp) = Source
                If doCancel Then Exit Function
                DoEvents
            Next
            Delete
    End Select
    ClearQueue 'Multiple levels of undo are stupid and probably buggy
    Undo = True
    End If
End Function
    

Public Function RecycleBinItems() As Long
    Dim RBinInfo As SHQUERYRBINFO
    'Const TwoGigabytes As Double = 2147483648#
    RBinInfo.cbSize = Len(RBinInfo)
    SHQueryRecycleBin vbNullString, RBinInfo
    If (RBinInfo.i64NumItems.LowPart And &H80000000) = &H80000000 Or RBinInfo.i64NumItems.HighPart > 0 Then
        RecycleBinItems = MaxLong
    Else
        RecycleBinItems = RBinInfo.i64NumItems.LowPart
    End If
End Function

Public Function RecycleBinBytes() As Long
    Dim RBinInfo As SHQUERYRBINFO
    'Const TwoGigabytes As Double = 2147483648#
    RBinInfo.cbSize = Len(RBinInfo)
    SHQueryRecycleBin vbNullString, RBinInfo
    If (RBinInfo.i64Size.LowPart And &H80000000) = &H80000000 Or RBinInfo.i64Size.HighPart > 0 Then
        RecycleBinBytes = MaxLong
    Else
        RecycleBinBytes = RBinInfo.i64Size.LowPart
    End If
End Function

Public Sub EmptyRecycleBin()
    SHEmptyRecycleBin UserControl.hwnd, vbNullString, SHERB_NOCONFIRMATION + SHERB_NOPROGRESSUI + SHERB_NOSOUND
    SHUpdateRecycleBinIcon
End Sub





Private Sub Scan(ByVal Filename As String, Optional Filter As String)
    Dim Temp As Long
    If Len(OriginalDir) = 0 Then OriginalDir = Filename
    
    Dirmain.Path = Filename
    For Temp = 0 To Dirmain.ListCount - 1
        AddFile Dirmain.List(Temp), True, False
    Next
    
    Filmain.Path = Filename
    If Right(Filename, 1) <> "\" Then Filename = Filename & "\"
    For Temp = 0 To Filmain.ListCount - 1
        If Len(Filter) = 0 Then
            AddFile Filename & Filmain.List(Temp)
        ElseIf IsLike(Filter, Filmain.List(Temp)) Then
            AddFile Filename & Filmain.List(Temp)
        End If
    Next
End Sub


Private Function AddFile(File As String, Optional IsADir As Boolean, Optional Recurse As Boolean, Optional Filter As String) As Long
    Dim Temp As Long, count As Long, Size As Long
    If InStr(File, "?") = 0 Then
        If IsADir Then
            count = DirCount
            AddFile = DirCount
            DirCount = DirCount + 1
            ReDim Preserve DirQueue(DirCount)
            DirQueue(DirCount - 1) = File
            'Debug.Print "Added dir: " & File
            
            If Recurse Then
                Scan File, Filter
                For Temp = count + 1 To DirCount - 1
                    Scan DirQueue(Temp), Filter
                Next
            End If
        Else
            AddFile = FileCount
            FileCount = FileCount + 1
            ReDim Preserve FileQueue(FileCount)
            FileQueue(FileCount - 1) = File
            
            If Len(File) > 3 And InStr(File, ":") > 0 Then
                Size = FileLen(File)
                If Size + FileSize <= MaxLong Then FileSize = FileSize + FileLen(File) Else FileSize = MaxLong
            End If
            'Debug.Print "Added file: " & File
        End If
    End If
End Function



Public Function IsADir(Filename As String) As Boolean
    On Error Resume Next
    If Len(Filename) > 0 Then IsADir = (GetAttr(Filename) And vbDirectory) = vbDirectory
End Function

Private Function CopyFile(ByVal from_file As String, ByVal to_file As String, Optional Cut As Boolean) As Boolean
    ' If a download is underway, complain.
    If Not isInProgress Then ' 0 Then Err.Raise fc_CopyInProgress, "FileCopier.CopyFile", "Copy already in progress"
        ' Do not copy if the file already exists.
        If Not FileExists(to_file) Then 'Err.Raise fc_FileExists, "FileCopier.CopyFile", "File already exists"
            isInProgress = True
            ' Save the from and to file names.
            m_FromFile = from_file
            m_ToFile = to_file

            If InStrRev(to_file, "\") > 3 Then MkPath Left(to_file, InStrRev(to_file, "\") - 1)
            ' Start the download.
            AsyncRead from_file, vbAsyncTypeFile, m_FromFile, vbAsyncReadForceUpdate
        
            If Cut Then DeleteFile from_file
            isInProgress = False
            CopyFile = True
            RaiseEvent FileComplete
        End If
    End If
End Function



Private Function MoveFile(ByVal from_file As String, ByVal to_file As String, Optional Cut As Boolean = True, Optional DoRAW As Boolean = True) As Boolean
    to_file = uniquefilename(to_file)
    RaiseEvent FileStarted(from_file)
    If isonsamedrive(from_file, to_file) Then
        If Cut Then
            Name from_file As to_file
        Else
            If DoRAW Then
                MoveFile = CopyFileRAW(from_file, to_file)
            Else
                MoveFile = CopyFile(from_file, to_file)
            End If
        End If
    Else
        If DoRAW Then
            MoveFile = CopyFileRAW(from_file, to_file, Cut)
        Else
            MoveFile = CopyFile(from_file, to_file, Cut)
        End If
    End If
End Function

Private Function isfilelocal(File As String) As Boolean
    isfilelocal = InStr(File, ":") = 2
End Function
Private Function isonsamedrive(ByVal from_file As String, ByVal to_file As String) As Boolean
    If isfilelocal(from_file) Then
        If isfilelocal(to_file) Then isonsamedrive = getdrive(from_file) = getdrive(to_file)
    End If
End Function
Private Function getdrive(File As String) As String
    getdrive = UCase(Left(File, 1))
End Function


Private Function ShellFile(hwnd As Long, ByVal File As String, Optional strOperation As String = "Open", Optional ShowAs As ShellExecuteConstants = SW_SHOWNORMAL)
    ShellFile = ShellExecute(hwnd, strOperation, File, vbNullString, App.Path, ShowAs)
End Function

Public Function RenameFile(Source As String, Destination As String) As Boolean
    On Error Resume Next
    Dim lpFileOp As SHFILEOPSTRUCT
    If InStr(Destination, "\") = 0 Then Destination = Left(Source, InStrRev(Source, "\")) & Destination
    With lpFileOp
        .fFlags = FOF_NOCONFIRMATION + FOF_NOERRORUI + FOF_SILENT + FOF_ALLOWUNDO
        'If Recycle Then .fFlags = .fFlags
        .hwnd = UserControl.hwnd
        .pFrom = Source & Chr(0) & Chr(0)
        .pTo = Destination & Chr(0) & Chr(0)
        .wFunc = FO_RENAME
    End With
    RenameFile = SHFileOperation(lpFileOp) = 0
End Function

Public Function DeleteFile(Path As String, Optional Recycle As Boolean = True) As Boolean
    Const FOF_NOERRORUI As Long = &H400
    On Error Resume Next
    RaiseEvent FileStarted(Path)
    Dim lpFileOp As SHFILEOPSTRUCT
    With lpFileOp
        .fFlags = FOF_NOCONFIRMATION + FOF_NOERRORUI + FOF_SILENT
        If Recycle Then .fFlags = .fFlags + FOF_ALLOWUNDO
        .hwnd = UserControl.hwnd
        .pFrom = Path & Chr(0) & Chr(0)
        .wFunc = FO_DELETE
    End With
    'Kill Path
    DeleteFile = SHFileOperation(lpFileOp) = 0
    RaiseEvent FileComplete
End Function


Private Function RemDir(Path As String) As Boolean
    On Error Resume Next
    RmDir Path
    RemDir = True
End Function

Public Sub MkPath(Path As String)
    Dim Temp As Long
    On Error Resume Next
    'temp = 3
    If Right(Path, 1) <> "\" Then Path = Path & "\"
    'Do Until temp = 1
    '    temp = InStr(temp, Path, "\") + 1
    '    If Not direxists(Left(Path, temp - 1)) Then MkDir Left(Path, temp - 1)
    'Loop
    
    Dim tempstr As String
    Temp = 3
    Do Until Temp = 0
        Temp = InStr(Temp + 1, Path, "\")
        If Temp > 0 Then
            If Not direxists(Left(Path, Temp - 1)) Then
                MkDir Left(Path, Temp - 1)
            End If
        End If
    Loop
End Sub

Public Function direxists(Directory As String) As Boolean
'Checks to see if a directory exists
    On Error Resume Next
    direxists = Len(Dir(Directory, vbDirectory + vbHidden)) > 0
End Function

Public Function FileExists(Filename As String) As Boolean
'Checks to see if a file exists
    On Error Resume Next
    FileExists = Len(Dir(Filename, vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem)) > 0
End Function

Private Function CreatePath(ByVal Path As String) As Boolean
    On Error Resume Next
    Dim i As Integer, s As String
    If direxists(Path) Then
        CreatePath = True
    Else
        If InStrRev(Path, "\") <> Len(Path) Then Path = Path + "\"
        Do
            i = InStr(i + 1, Path, "\")
            If i = 0 Then Exit Do
            s = Left$(Path, i - 1)
            If Not direxists(s) Then MkDir s
        Loop Until i = Len(Path)
        CreatePath = direxists(Path)
    End If
End Function

Private Sub DELTREE(spathname As String)
    On Error Resume Next
    Dim sFileName As String, dSize As Double, asFileName() As String, i As Long
    If StrComp(Right$(spathname, 1), "\", vbBinaryCompare) <> 0 Then spathname = spathname & "\"
    sFileName = Dir$(spathname, vbDirectory + vbHidden + vbSystem + vbReadOnly)
    Do While Len(sFileName) > 0
        If StrComp(sFileName, ".", vbBinaryCompare) <> 0 And StrComp(sFileName, "..", vbBinaryCompare) <> 0 Then
            ReDim Preserve asFileName(i)
            asFileName(i) = spathname & sFileName
            i = i + 1
        End If
        sFileName = Dir
    Loop
    If i > 0 Then
        For i = 0 To UBound(asFileName)
            If (GetAttr(asFileName(i)) And vbDirectory) = vbDirectory Then
                DELTREE (asFileName(i))
            Else
                Kill FileLen(asFileName(i))
            End If
        Next
        RmDir spathname
    End If
End Sub




Public Function CopyFileRAW(Source As String, Destination As String, Optional ByVal BufferLen As Long = DefaultBuffer, Optional UpdateInterval As Long = 10, Optional Cut As Boolean) As Boolean
    Dim SRCfile As Long, DESTfile As Long, buffer() As Byte, LENleft As Long, cPos As Long
    Dim StartTime As Date, TotalLen As Long, ETA As Long, Diff As Long, BytesPerSecond As Double, Update As Long
    On Error Resume Next
    
    If BufferLen = 0 Then BufferLen = DefaultBuffer
    ReDim buffer(BufferLen)
    SRCfile = FreeFile
    Open Source For Binary As #SRCfile
    
    If FileExists(Destination) Then DeleteFile Destination, True
    DESTfile = FreeFile
    
    MkPath Left(Destination, InStrRev(Destination, "\"))
    Open Destination For Binary As #DESTfile
    
    LENleft = FileLen(Source)
    TotalLen = LENleft
    cPos = 1
    StartTime = Now
    
    Do While LENleft > 0
        If LENleft < BufferLen Then
            BufferLen = LENleft
            ReDim buffer(BufferLen)
        End If
    
        Get #SRCfile, cPos, buffer
        Put #DESTfile, cPos, buffer
        
        cPos = cPos + BufferLen
        LENleft = LENleft - BufferLen
        
        Update = Update + 1
        If Update >= UpdateInterval Then
            Diff = DateDiff("s", StartTime, Now)
            If Diff > 0 Then
                BytesPerSecond = cPos / Diff
                ETA = LENleft / BytesPerSecond
                RaiseEvent FileProgressReport(cPos, TotalLen, ETA)
                Update = 0
            End If
        End If
        DoEvents
        
        If doCancel Then
            If LENleft > 0 Then
                LENleft = 0
                Source = Destination
                Cut = True
                GoTo exitsub
            End If
        End If
    Loop
    
exitsub:
    Close #SRCfile
    Close #DESTfile
    
    If Cut Then DeleteFile Source
    
    RaiseEvent FileComplete
    CopyFileRAW = True
End Function

Private Function ReplaceNumbers(ByVal Text As String, Optional Numbers As Long, Optional StartAt As Long) As String
    Dim Temp As Long, tempstr As String, Abort As Boolean, char As String, Found As Long, Start As String
    tempstr = Left(Text, InStrRev(Text, "\"))
    Text = Right(Text, Len(Text) - InStrRev(Text, "\"))
    Abort = Numbers > 0
    For Temp = 1 To Len(Text)
        char = Mid(Text, Temp, 1)
        If char >= "0" And char <= "9" Then
            Found = Found + 1
            If Found <= Numbers Or Numbers = 0 Then tempstr = tempstr & "#"
            Start = Start & char
        Else
            tempstr = tempstr & char
        End If
    Next
    If Len(Start) > 0 Then StartAt = Val(Start) Else StartAt = 1
    ReplaceNumbers = tempstr
End Function

Public Function uniquefilename(Filename As String, Optional ByVal Pattern = " (#)") As String
    Dim temp1 As String, temp2 As String, temp3 As Long, Found As Boolean, NewFile As String, Dir As Boolean
    uniquefilename = Filename
    Dir = direxists(Filename)
    If FileExists(Filename) Or Dir Then
        Dim count As Long
        count = 1
        temp3 = InStrRev(Filename, ".")
        temp1 = Filename
        
        If Len(Pattern) = 0 Then
            Pattern = ReplaceNumbers(Filename, 1, count)
            If InStr(Pattern, "#") > 0 Then
                temp1 = Left(Pattern, InStrRev(Pattern, "#") - 1)
                temp2 = Right(Pattern, Len(Pattern) - InStrRev(Pattern, "#"))
            Else
                temp1 = Left(Filename, temp3 - 1)
                temp2 = Right(Filename, Len(Filename) - temp3 + 1)
            End If
            Pattern = "#"
        Else
            If temp3 > 0 And Not Dir Then
                temp1 = Left(Filename, temp3 - 1)
                temp2 = Right(Filename, Len(Filename) - temp3 + 1)
            End If
        End If
        
        Do Until Found
            NewFile = temp1 & Replace(Pattern, "#", count) & temp2
            If Dir Then
                Found = Not direxists(NewFile)
            Else
                Found = Not FileExists(NewFile)
            End If
            count = count + 1
        Loop
        
        'If Dir Then
        '    Do Until Not direxists(temp1 & " (" & count & ")" & temp2)
        '        count = count + 1
        '    Loop
        'Else
        '    Do Until Not FileExists(temp1 & " (" & count & ")" & temp2)
        '        count = count + 1
        '    Loop
        'End If
        uniquefilename = NewFile
    End If
End Function
