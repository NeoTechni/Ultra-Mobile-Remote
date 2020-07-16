Attribute VB_Name = "SystemAPI"
Option Explicit
'install font
'Declare Function WriteProfileString Lib "Kernel" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal lpString As String) As Integer
Declare Function WriteProfileString Lib "Kernel32.dll" Alias "WriteProfileStringA" (ByVal lpszSection As String, ByVal lpszKeyName As String, ByVal lpszString As String) As Long
'Declare Function CreateScalableFontResource% Lib "GDI32" (ByVal fHidden%, ByVal lpszResourceFile$, ByVal lpszFontFile$, ByVal lpszCurrentPath$)
Declare Function CreateScalableFontResource Lib "gdi32.dll" Alias "CreateScalableFontResourceA" (ByVal fdwHidden As Long, ByVal lpszFontRes As String, ByVal lpszFontFile As String, ByVal lpszCurrentPath As String) As Long
Declare Function AddFontResource Lib "gdi32.dll" Alias "AddFontResourceA" (ByVal lpszFileName As String) As Long
'Declare Function AddFontResource Lib "gdi32.dll" (ByVal lpFilename As Any) As Integer
'Declare Function SendMessage Lib "User" (ByVal hwnd As Integer, ByVal wMsg As Integer, ByVal wParam As Integer, lParam As Any) As Long
Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

Public Declare Function GetWindowsDirectory Lib "Kernel32.dll" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long



Public Const Kilobyte = 1024
Public Const Megabyte = 1048576
Public Const Gigabyte = 1073741824

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Declare Function FileTimeToLocalFileTime Lib "kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function GetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFilename As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function SystemTimeToFileTime Lib "kernel32" (lpSystemTime As SYSTEMTIME, lpFileTime As FILETIME) As Long
Private Declare Function LocalFileTimeToFileTime Lib "kernel32" (lpLocalFileTime As FILETIME, lpFileTime As FILETIME) As Long
Private Declare Function SetFileTime Lib "kernel32" (ByVal hFile As Long, lpCreationTime As FILETIME, lpLastAccessTime As FILETIME, lpLastWriteTime As FILETIME) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Const GENERIC_READ As Long = &H80000000
Private Const GENERIC_WRITE = &H40000000
Private Const OPEN_EXISTING = 3
Private Const FILE_SHARE_READ = &H1
Private Const FILE_SHARE_WRITE = &H2

Private Declare Function SHGetDiskFreeSpace Lib "Shell32" Alias "SHGetDiskFreeSpaceA" (ByVal pszVolume As String, pqwFreeCaller As Currency, pqwTot As Currency, pqwFree As Currency) As Long
'Private Declare Function GetLogicalDriveStrings Lib "kernel32" Alias "GetLogicalDriveStringsA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetDriveType Lib "kernel32" Alias "GetDriveTypeA" (ByVal nDrive As String) As Long
'Private Declare Function GetVolumeInformation Lib "kernel32" Alias "GetVolumeInformationA" (ByVal lpRootPathName As String, ByVal lpVolumeNameBuffer As String, ByVal nVolumeNameSize As Long, lpVolumeSerialNumber As Long, lpMaximumComponentLength As Long, lpFileSystemFlags As Long, ByVal lpFileSystemNameBuffer As String, ByVal nFileSystemNameSize As Long) As Long
Private Const DRIVE_HARD = 1, DRIVE_REMOVABLE = 2, DRIVE_FIXED = 3, DRIVE_REMOTE = 4, DRIVE_CDROM = 5, DRIVE_RAMDISK = 6
'Private Declare Function GetDiskFreeSpace Lib "kernel32" Alias "GetDiskFreeSpaceA" (ByVal lpRootPathName As String, lpSectorsPerCluster As Long, lpBytesPerSector As Long, lpNumberOfFreeClusters As Long, lpTotalNumberOfClusters As Long) As Long
Private Declare Function SHGetFolderPath Lib "shell32.dll" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwFlags As Long, ByVal lpszPath As String) As Long

Public Enum FolderType
    CSIDL_APPDATA = &H1A
    'NT only!
End Enum

Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long

Public Enum ShellExecuteConstants
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


Public Function ShellFile(hwnd As Long, ByVal File As String, Optional strOperation As String = "Open", Optional ShowAs As ShellExecuteConstants = SW_SHOWNORMAL)
    ShellFile = ShellExecute(hwnd, strOperation, File, vbNullString, App.Path, ShowAs)
End Function

Public Function SystemFolder(Owner As Form, FolderID As FolderType) As String
    Const SHGFP_TYPE_CURRENT = 0, SHGFP_TYPE_DEFAULT = 1
    Dim Path As String
    SHGetFolderPath Owner.hwnd, FolderID, 0, SHGFP_TYPE_CURRENT, Path
    SystemFolder = Path
End Function

Public Function GetPercentFree(DriveRoot As String) As Double
    Dim FreeCaller As Currency, Tot As Currency, Free As Currency
    SHGetDiskFreeSpace Left(DriveRoot, 1) & ":\", FreeCaller, Tot, Free
    GetPercentFree = Free / Tot
End Function
Public Function GetPercentFull(DriveRoot As String) As Double
    GetPercentFull = 1 - GetPercentFree(DriveRoot)
End Function

Public Function GetTotalSpaceGigaBytes(DriveRoot As String) As Double
    Dim FreeCaller As Currency, Tot As Currency, Free As Currency
    SHGetDiskFreeSpace Left(DriveRoot, 1) & ":\", FreeCaller, Tot, Free
    GetTotalSpaceGigaBytes = Tot * 10000 / Gigabyte
End Function

Public Function GetFreeSpaceGigaBytes(DriveRoot As String) As Double
    Dim FreeCaller As Currency, Tot As Currency, Free As Currency
    SHGetDiskFreeSpace Left(DriveRoot, 1) & ":\", FreeCaller, Tot, Free
    GetFreeSpaceGigaBytes = Free * 10000 / Gigabyte
End Function

Public Function FileTitle(Path As String, Optional RemoveExtention As Boolean) As String
    Dim tempstr As String
    tempstr = Right(Path, Len(Path) - InStrRev(Path, "\"))
    If RemoveExtention Then
        If InStr(tempstr, ".") > 0 Then tempstr = Left(tempstr, InStrRev(tempstr, ".") - 1)
    End If
    FileTitle = tempstr
End Function
Public Function ChkFile(Path As String, File As String) As String
    ChkFile = Replace(Path & "\" & File, "\\", "\")
End Function

Public Function CopyFolder(Dirbox As DirListBox, FILEbox As FileListBox, Source As String, Destination As String, Optional RecurseSubs As Boolean, Optional Delete As Boolean) As Boolean
    Dim temp As Long, tempstr As String
    'On Error Resume Next
    tempstr = Dirbox.Path
    If Not Delete Then CreatePath Destination
    
    If RecurseSubs Then
        Dirbox.Path = Source 'Replace(Replace(Source, "/", "-"), ":", "-")
        For temp = 0 To Dirbox.ListCount - 1
            CopyFolder Dirbox, FILEbox, Dirbox.List(temp), Destination & FileTitle(Dirbox.List(temp)), True
        Next
        Dirbox.Path = tempstr
    End If
    
    FILEbox.Path = Source
    For temp = 0 To FILEbox.ListCount - 1
        If Delete Then
            DeleteFile ChkFile(Source, FILEbox.List(temp))
        Else
            CopyFile ChkFile(Source, FILEbox.List(temp)), ChkFile(Destination, FILEbox.List(temp))
        End If
    Next
    If Delete Then CopyFolder = RemDir(Source) Else CopyFolder = True
End Function

Public Sub DeleteFile(Path As String)
    On Error Resume Next
    Kill Path
End Sub

Public Function RemDir(Path As String)
    On Error Resume Next
    RmDir Path
    RemDir = Not direxists(Path)
End Function

Public Function CopyFile(Source As String, Destination As String, Optional Cut As Boolean) As Boolean
    On Error Resume Next
    FileCopy Source, Destination
    CopyFile = FileExists(Destination)
    If Cut Then DeleteFile Source
End Function
Public Sub MkPath(Path As String)
    Dim temp As Long
    temp = 3
    Do Until temp = 1
        temp = InStr(temp, Path, "\") + 1
        If Not direxists(Left(Path, temp - 1)) Then MkDir Left(Path, temp - 1)
    Loop
End Sub

Public Function direxists(Directory As String) As Boolean
'Checks to see if a directory exists
On Error Resume Next
direxists = Len(DIR(Directory, vbDirectory + vbHidden)) > 0
End Function

Public Function FileExists(Filename As String) As Boolean
'Checks to see if a file exists
On Error Resume Next
    FileExists = Len(DIR(Filename, vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem)) > 0
End Function

Public Function CreatePath(ByVal Path As String) As Boolean
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
Public Sub DELTREE(spathname As String)
    On Error Resume Next
    Dim sFileName As String, dSize As Double, asFileName() As String, i As Long
    If StrComp(Right$(spathname, 1), "\", vbBinaryCompare) <> 0 Then spathname = spathname & "\"
    sFileName = DIR$(spathname, vbDirectory + vbHidden + vbSystem + vbReadOnly)
    Do While Len(sFileName) > 0
        If StrComp(sFileName, ".", vbBinaryCompare) <> 0 And StrComp(sFileName, "..", vbBinaryCompare) <> 0 Then
            ReDim Preserve asFileName(i)
            asFileName(i) = spathname & sFileName
            i = i + 1
        End If
        sFileName = DIR
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

Public Function DriveType(driveletter As String, Optional Hard As String = "Hard Drive", Optional Floppy As String = "Floppy Drive", Optional Network As String = "Network Drive", Optional Optical As String = "Optical Drive", Optional RAM As String = "RAM Drive", Optional Unknown As String = "Unknown Drive Type") As String
Select Case GetDriveType(Left(driveletter, 1) & ":\")
    Case 3:        DriveType = Hard
    Case 2:        DriveType = Floppy
    Case 4:        DriveType = Network
    Case 5:        DriveType = Optical
    Case 6:        DriveType = RAM
    Case Else:     DriveType = Unknown
End Select
End Function



Private Function GetFileTimes(ByVal file_name As String, ByRef date_created As Date, ByRef date_accessed As Date, ByRef date_written As Date, ByVal local_time As Boolean, Optional GetCreation As Boolean = True, Optional GetAccessed As Boolean = True, Optional GetWritten As Boolean = True) As Boolean
    Dim file_handle As Long
    Dim creation_time As FILETIME
    Dim access_time As FILETIME
    Dim write_time As FILETIME
    Dim file_time As FILETIME

    ' Open the file.
    file_handle = CreateFile(file_name, GENERIC_READ, FILE_SHARE_READ Or FILE_SHARE_WRITE, 0&, OPEN_EXISTING, 0&, 0&)
    If file_handle = 0 Then
        GetFileTimes = True
        Exit Function
    End If

    ' Get the times.
    If GetFileTime(file_handle, creation_time, access_time, write_time) = 0 Then
        GetFileTimes = True
        Exit Function
    End If

    ' Close the file.
    If CloseHandle(file_handle) = 0 Then
        GetFileTimes = True
        Exit Function
    End If

    ' See if we should convert to the local
    ' file system time.
    If local_time Then
        ' Convert to local file system time.
        If GetCreation Then
            FileTimeToLocalFileTime creation_time, file_time
            creation_time = file_time
        End If
        
        If GetAccessed Then
            FileTimeToLocalFileTime access_time, file_time
            access_time = file_time
        End If
        
        If GetWritten Then
            FileTimeToLocalFileTime write_time, file_time
            write_time = file_time
        End If
    End If

    ' Convert into dates.
    If GetCreation Then date_created = FileTimeToDate(creation_time)
    If GetAccessed Then date_accessed = FileTimeToDate(access_time)
    If GetWritten Then date_written = FileTimeToDate(write_time)
End Function

' Convert the FILETIME structure into a Date.
Private Function FileTimeToDate(ft As FILETIME) As Date
    ' FILETIME units are 100s of nanoseconds.
    Const TICKS_PER_SECOND = 10000000

    Dim lo_time As Double
    Dim hi_time As Double
    Dim seconds As Double
    Dim hours As Double
    Dim the_date As Date

    ' Get the low order data.
    If ft.dwLowDateTime < 0 Then
        lo_time = 2 ^ 31 + (ft.dwLowDateTime And &H7FFFFFFF)
    Else
        lo_time = ft.dwLowDateTime
    End If

    ' Get the high order data.
    If ft.dwHighDateTime < 0 Then
        hi_time = 2 ^ 31 + (ft.dwHighDateTime And &H7FFFFFFF)
    Else
        hi_time = ft.dwHighDateTime
    End If

    ' Combine them and turn the result into hours.
    seconds = (lo_time + 2 ^ 32 * hi_time) / TICKS_PER_SECOND
    hours = CLng(seconds / 3600)
    seconds = seconds - hours * 3600

    ' Make the date.
    the_date = DateAdd("h", hours, "1/1/1601 0:00 AM")
    the_date = DateAdd("s", seconds, the_date)
    FileTimeToDate = the_date
End Function

Public Function FileCreationDate(Path As String) As Date
    Dim Created As Date, Accessed As Date, Written As Date
    GetFileTimes Path, Created, Accessed, Written, True, True, False, False
    FileCreationDate = Created
End Function
Public Function FileAccessedDate(Path As String) As Date
    Dim Created As Date, Accessed As Date, Written As Date
    GetFileTimes Path, Created, Accessed, Written, True, False, True, False
    FileAccessedDate = Accessed
End Function

'Star Date (galactic standard time) relative from the year 2323 (invention of warp drive was way before this, so I have no idea why this year)
Public Function StarDate(theTime As Date, Optional DigitsAfterDecimal As Long = -1) As Double
    Dim Year As Long, Day As Long, hour As Double
    Year = (Format(theTime, "YYYY") - 2323) * 1000
    Day = DatePart("y", theTime) / DaysInYear(Format(theTime, "YYYY")) * 1000
    hour = (Format(theTime, "HH") * 3600 + Format(theTime, "nn") * 60 + Format(theTime, "ss")) / 86400
    If DigitsAfterDecimal > -1 Then hour = Round(hour, DigitsAfterDecimal)
    StarDate = Year + Day + hour
End Function

'Date functions
Public Function DaysInYear(Year As Long) As Long
    DaysInYear = DatePart("y", DateSerial(Year, 12, 31))
End Function
Public Function DaysInMonth(ByVal month As Long, Year As Long) As Long
    DaysInMonth = Val(Format(DateSerial(Year, month + 1, 0), "dd"))
End Function
Public Function GetDate(ByVal Day As Long, Optional Year As Long = 2001) As Date
    GetDate = DateAdd("d", Day - 1, DateSerial(Year, 1, 1))
End Function




Public Function GetExtention(ByVal Filename As String) As String
    If InStr(Filename, ".") Then GetExtention = Right(Filename, Len(Filename) - InStrRev(Filename, "."))
End Function
Public Function FindExtention(List, Count As Long, Extention As String) As Long
    Dim temp As Long
    FindExtention = -1
    For temp = 0 To Count - 1
        If StrComp(Extention, List(temp), vbTextCompare) = 0 Then
            FindExtention = temp
            Exit For
        End If
    Next
End Function
Public Function EnumExtentions(FILEbox As FileListBox, List) As Long
    Dim temp As Long, Count As Long, temp2 As Long, tempstr As String
    For temp = 0 To FILEbox.ListCount - 1
        tempstr = GetExtention(FILEbox.List(temp))
        temp2 = FindExtention(List, Count, tempstr)
        If temp2 = -1 Then
            Count = Count + 1
            ReDim Preserve List(Count)
            List(Count - 1) = tempstr
        End If
    Next
    EnumExtentions = Count
End Function
Public Sub SortExtentions(List, Count As Long)
    Dim temp As Long, temp2 As Long, tempstr As String
    For temp = 1 To Count - 1
        For temp2 = temp - 1 To 0 Step -1
            If StrComp(List(temp2 + 1), List(temp2), vbTextCompare) = -1 Then
                tempstr = List(temp2)
                List(temp2) = List(temp2 + 1)
                List(temp2 + 1) = tempstr
            Else
                Exit For
            End If
        Next
    Next
End Sub

Public Function DebugExtentions(List, Count As Long) As String
    Dim temp As Long, tempstr As String
    If Count > 0 Then
        tempstr = List(0)
        For temp = 1 To Count - 1
            tempstr = tempstr & ", " & List(temp)
        Next
        DebugExtentions = tempstr
    End If
End Function

'not needed, too slow
Public Function EnumFilesExt(FILEbox As FileListBox, List, Extention As String) As Long
    Dim temp As Long, Count As Long
    For temp = 0 To FILEbox.ListCount - 1
        If StrComp(GetExtention(FILEbox.List(temp)), Extention, vbTextCompare) = 0 Then
            Count = Count + 1
            ReDim Preserve List(Count)
            List(Count - 1) = Replace(FILEbox.Path & "\" & FILEbox.List(temp), "\\", "\")
        End If
    Next
    EnumFilesExt = Count
End Function

Public Function FindExtIndex(Filename As String, ExtList, ExtCount As Long) As Long
    Dim Extention As String, temp As Long
    Extention = GetExtention(Filename)
    FindExtIndex = -1
    For temp = 0 To ExtCount - 1
        If StrComp(Extention, ExtList(temp), vbTextCompare) = 0 Then
            FindExtIndex = temp
            Exit For
        End If
    Next
End Function

Private Sub SwitchItem(List, ItemA As Long, ItemB As Long)
    Const UpperBound As Long = 3
    Dim tempstr(0 To UpperBound) As String, temp As Long
    For temp = 0 To UpperBound
        tempstr(temp) = List(temp, ItemA)
        List(temp, ItemA) = List(temp, ItemB)
        List(temp, ItemB) = tempstr(temp)
    Next
End Sub

'Sortby: 0 = not sorted, 1 = sort by type, 2 = sort by size, 3 = sort by file creation date, 4 = sort by modified date, 5 = accessed date
Public Function EnumSortedFiles(FILEbox As FileListBox, List, Optional SortBy As Long = 1) As Long
    Dim temp As Long, temp2 As Long, temp3 As Long, temp4 As Long
    Dim Count As Long, ExtList() As String, ExtCount As Long
    Dim ExtCounts() As Long, cCount As Long, tempfile As String, tempExt As Long
    
    ExtCount = EnumExtentions(FILEbox, ExtList)
    SortExtentions ExtList, ExtCount
    ReDim ExtCounts(ExtCount)
    
    Count = FILEbox.ListCount
    ReDim List(0 To 3, Count)
    For temp = 0 To FILEbox.ListCount - 1
        If InStr(FILEbox.List(temp), "?") > 0 Then
            If Count > 0 Then
                Count = Count - 1
                ReDim Preserve List(0 To 3, Count)
            End If
        Else
            List(0, temp4) = Replace(FILEbox.Path & "\" & FILEbox.List(temp), "\\", "\")
            List(1, temp4) = FindExtIndex(FILEbox.List(temp), ExtList, ExtCount)
            ExtCounts(List(1, temp4)) = ExtCounts(List(1, temp4)) + 1
            List(2, temp4) = FileLen(List(0, temp4))
        
            Select Case SortBy
                'Case 3: List(3, temp) = CStr(FileCreationDate( List(0, temp)))
                'Case 4: List(3, temp) = CStr(FileDateTime(List(0, temp))) 'modified date
                'Case 5: List(3, temp) = CStr(FileAccessedDate(List(0, temp)))
            End Select
        
            temp4 = temp4 + 1
        End If
    Next
    
Select Case SortBy ' = 1 Then
    Case 0 'GNDN
    Case 1 'type
        For temp = 0 To ExtCount - 1
            For temp2 = 0 To Count - 1
                If List(1, temp2) = temp Then
                    'is a file matching current extention type
                    For temp3 = temp2 - 1 To cCount Step -1
                        'tempFile = List(0, temp3)
                        'tempExt = List(1, temp3)
                    
                        'List(0, temp3) = List(0, temp3 + 1)
                        'List(1, temp3) = List(1, temp3 + 1)
                    
                        'List(0, temp3 + 1) = tempFile
                        'List(1, temp3 + 1) = tempExt
                        
                        SwitchItem List, temp3, temp3 + 1
                    Next
                    cCount = cCount + 1
                End If
            Next
        Next
    Case 2 'size
    Case 3, 4, 5 'sort by dates
End Select
    'For temp = 0 To Count - 1
    '    Debug.Print List(0, temp)
    'Next
    EnumSortedFiles = Count
End Function



Public Function SizeToText(Size As Long, Optional Bytes As String = "B", Optional KB As String = "K", Optional MB As String = "M", Optional GB As String = "G", Optional DigitsAfterDecimal As Long = 2) As String
    Select Case Size 'FileLen(File)
        Case 0 To 1023:             SizeToText = Size & Bytes
        Case 1024 To 1048575:       SizeToText = Round(Size / 1024, DigitsAfterDecimal) & KB
        Case 1048576 To 1073741823: SizeToText = Round(Size / 1048576, DigitsAfterDecimal) & MB
        Case Else:                  SizeToText = Round(Size / 1073741824, DigitsAfterDecimal) & GB
    End Select
End Function
Public Function SizeToLCAR(Size As Long) As Long
    Dim DigitsAfterDecimal As Long, DivideBy As Long, Value As Single, sValue As String, y As Long, x As Long, Dec As Long
    Dim DivideBy2 As Long, Value2 As Long
    
    Select Case Size
        Case 0 To 9:                    x = 0: DivideBy2 = 1
        Case 10 To 99:                  x = 1: DivideBy2 = 10
        Case 100 To 999:                x = 2: DivideBy2 = 100
        Case 1000 To 9999:              x = 3: DivideBy2 = 1000
        Case 10000 To 99999:            x = 4: DivideBy2 = 10000
        Case 100000 To 999999:          x = 5: DivideBy2 = 100000
        Case 1000000 To 9999999:        x = 6: DivideBy2 = 1000000
        Case 10000000 To 99999999:      x = 7: DivideBy2 = 10000000
        Case 100000000 To 999999999:    x = 8: DivideBy2 = 100000000
        Case Else:                      x = 9: DivideBy2 = 1000000000
    End Select
    
    Value = Size / DivideBy2
    Select Case Value
        Case 0 To 99:                   DigitsAfterDecimal = 2
        Case 100 To 999:                DigitsAfterDecimal = 1
            x = x + 1
            Value = Value / 10
        Case Else
            x = x + 3
            Value = Value / 1000
    End Select
    
    Value = Round(Value, DigitsAfterDecimal)
    sValue = CStr(Value)
    Dec = InStr(sValue, ".")
    If Dec > 0 Then
        Dec = Len(sValue) - Dec
        Value = Value * (10 ^ Dec)
        x = x - Dec
        sValue = CStr(Value)
    End If
    
    'Debug.Print sValue & ", " & X & ", " & Value * (10 ^ X)
    'Debug.Print sValue & ", " & Y & ", " & Value * (2 ^ Y)
    
    SizeToLCAR = Value * 10 + x
End Function


Public Function Kilobytes(Size As Long) As Double
    Kilobytes = Size / 1024
End Function
Public Function Gigabytes(Size As Long) As Double
    Gigabytes = Size / 1048576
End Function
Public Function Terabytes(Size As Long) As Double
    Terabytes = Size / 1073741824
End Function



'install font

Public Function WinDir(Optional ByVal SysDir As Boolean = False) As String
    Dim t As String * 255
    Dim i As Long
    If SysDir Then
        i = GetSystemDirectory(t, Len(t))
    Else
        i = GetWindowsDirectory(t, Len(t))
    End If
    WinDir = Left(t, i)
End Function

Public Function GetFontName(FileNameTTF As String) As String

   Dim hFile As Integer
   Dim buffer As String
   Dim FontName As String
   Dim TempName As String
   Dim iPos As Integer
   
  'Build name for new resource file in a temporary file, and call the API.
   TempName = App.Path & "\~TEMP.FOT"
   If CreateScalableFontResource(1, TempName, FileNameTTF, vbNullString) Then
      
     'The name sits behind the text "FONTRES:"
      hFile = FreeFile
      Open TempName For Binary Access Read As hFile

         buffer = SPACE(LOF(hFile))
         Get hFile, , buffer
         iPos = InStr(buffer, "FONTRES:") + 8
         FontName = Mid(buffer, iPos, InStr(iPos, buffer, vbNullChar) - iPos)
      Close hFile
      Kill TempName
    End If
  'Return the font name
   GetFontName = FontName
End Function

'Color code
Public Function Red(Color As Long) As Byte
    Red = Color Mod 256
End Function

Public Function Green(Color As Long) As Byte
    Green = ((Color And &HFF00) / 256) Mod 256
End Function

Public Function Blue(Color As Long) As Byte
    Blue = (Color And &HFF0000) / 65536
End Function

'Alpha blend functions
Private Function SysToLNG(ByVal lColor As Long) As Long
'Special thanks to redbird77 for this code and realizing what the bug was
' If hi-bit is set, then it is a system color.
If (lColor And &H80000000) Then SysToLNG = GetSysColor(lColor And &HFFFFFF) Else SysToLNG = lColor
End Function

Public Function HTMLColor(ByVal RGB As Long, Optional No As Boolean) As String
    RGB = SysToLNG(RGB)
    If No Then
        HTMLColor = Red(RGB) & "," & Green(RGB) & "," & Blue(RGB)
    Else
        HTMLColor = "#" & Format(Hex(Red(RGB)), "00") & Format(Hex(Green(RGB)), "00") & Format(Hex(Blue(RGB)), "00")
    End If
End Function
