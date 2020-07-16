Attribute VB_Name = "IShellFolderAPI"
Option Explicit
'
' Brad Martinez  http://www.mvps.org/ccrp
'
' Code was written in and formatted for 8pt MS San Serif
'
' ===========================================================
' Note that "IShellFolder Extended Type Library v1.1" (ISHF_Ex.tlb) included
' with this project, must be present, correctly registered on your system, and
' referenced by this project.
' ===========================================================

' Defined as an HRESULT that corresponds to S_OK.
Public Const NOERROR = 0

' Retrieves the IShellFolder interface for the desktop folder.
' Returns NOERROR if successful or an OLE-defined error result otherwise.
Declare Function SHGetDesktopFolder Lib "shell32" (ppshf As IShellFolder) As Long

' Retrieves a pointer to the shell's IMalloc interface.
' Returns NOERROR if successful or or E_FAIL otherwise.
Declare Function SHGetMalloc Lib "shell32" (ppMalloc As IMalloc) As Long
'
' ===========================================================
' item ID (pidl) structs, just for reference...
'
'' item identifier (relative pidl), allocated by the shell
'Public Type SHITEMID
'  cb As Integer        ' size of struct, including cb itself
'  abID(0) As Byte    ' variable length item identifier
'End Type
'
'' fully qualified pidl
'Public Type ITEMIDLIST
'  mkid As SHITEMID  ' list of item identifers, packed into SHITEMID.abID
'End Type
'

' Returns a reference to the IMalloc interface.

Public Function isMalloc() As IMalloc
  Static im As IMalloc
  
  ' SHGetMalloc should just get called once as the 'im'
  ' variable stays in scope while the project is running...
  If (im Is Nothing) Then Call SHGetMalloc(im)
  
  Set isMalloc = im

End Function

' Returns a reference to the desktop's IShellFolder interface.

Public Function isfDesktop() As IShellFolder
  Static isf As IShellFolder
  
  ' SHGetDesktopFolder should just get called once as the 'isf'
  ' variable stays in scope while the project is running...
  If (isf Is Nothing) Then Call SHGetDesktopFolder(isf)
  
  Set isfDesktop = isf

End Function

' Returns the IShellFolder's interface ID.

Public Function IID_IShellFolder() As Guid
  Static iid As Guid
  
  ' Fill the IShellFolder interface ID, {000214E6-000-000-C000-000000046}
  If (iid.Data1 = 0) Then
    iid.Data1 = &H214E6
    iid.Data4(0) = &HC0
    iid.Data4(7) = &H46
  End If
  
  IID_IShellFolder = iid
  
End Function

' =====================================================================
' pidl attributes

' Determines if the specified pidl is the desktop folder's pidl.
' Returns True if the pidl is the desktop's pidl, returns False otherwise.

' The desktop pidl is only a single item ID whose value is 0 (the 2 byte
' zero-terminator, i.e. SHITEMID.abID is empty). Direct descendents of
' the desktop (My Computer, Network Neighborhood) are absolute pidls
' (relative to the desktop) also with a single item ID, but contain values
' (SHITEMID.abID > 0). Drive folders have 2 item IDs, children of drive
' folders have 3 item IDs, etc. All other single item ID pidls are relative to
' the shell folder in which they reside (just like a relative path).

Public Function IsDesktopPIDL(pidl As Long) As Boolean
  
  ' The GetItemIDSize() call will also return 0 if pidl = 0
  If pidl Then IsDesktopPIDL = (GetItemIDSize(pidl) = 0)

End Function

' Returns the size in bytes of the first item ID in a pidl.
' Returns 0 if the pidl is the desktop's pidl or is the last
' item ID in the pidl (the zero terminator), or is invalid.

Public Function GetItemIDSize(ByVal pidl As Long) As Integer
  
  ' If we try to access memory at address 0 (NULL), then it's bye-bye...
  If pidl Then MoveMemory GetItemIDSize, ByVal pidl, 2

End Function

' Returns a pointer to the next item ID in a pidl.
' Returns 0 if the next item ID is the pidl's zero value terminating 2 bytes.

Public Function GetNextItemID(ByVal pidl As Long) As Long
  Dim cb As Integer   ' SHITEMID.cb, 2 bytes
  
  cb = GetItemIDSize(pidl)
  ' Make sure it's not the zero value terminator.
  If cb Then GetNextItemID = pidl + cb

End Function

' If successful, returns the size in bytes of the memory occcupied by a pidl,
' including it's 2 byte zero terminator. Returns 0 otherwise.

Public Function GetPIDLSize(ByVal pidl As Long) As Integer
  Dim cb As Integer
  ' Error handle in case we get a bad pidl and overflow cb.
  ' (most item IDs are roughly 20 bytes in size, and since an item ID represents
  ' a folder, a pidl can never exceed 260 folders, or 5200 bytes).
  On Error GoTo Out
  
  If pidl Then
    Do While pidl
      cb = cb + GetItemIDSize(pidl)
      pidl = GetNextItemID(pidl)
    Loop
    ' Add 2 bytes for the zero terminating item ID
    GetPIDLSize = cb + 2
  End If
  
Out:
End Function

' =================================================================
' displayname

' Returns a folder's displayname

'   isfParent - pidl's parent IShellFolder
'   pidlRel    - child object's relative pidl we're getting the name of
'   uFlags    - specifies the type of name to retrieve

Public Function GetFolderDisplayName(isfParent As IShellFolder, pidlRel As Long, uFlags As ESHGNO) As String
  Dim lpStr As STRRET   ' struct filled
  Dim lpsz As Long          ' temp string pointer
  Dim uOffset As Long     ' offset to the string pointer
  Dim sName As String     ' return string
  On Error GoTo Out
  
  ' returns 0x80004001(E_NOTIMPL) for non SHGDN_FORPARSING net absolute pidls
  If (isfParent.GetDisplayNameOf(pidlRel, uFlags, lpStr) = NOERROR) Then
    Select Case (lpStr.uType)

      ' The 1st UINT (Long) of the array points to a Unicode
      ' str which *should* be allocated & freed (?).
      Case STRRET_WSTR:
        MoveMemory lpsz, lpStr.CStr(0), 4
        sName = String$(MAX_PATH, 0)
        Call WideCharToMultiByte(CP_ACP, 0, ByVal lpsz, -1, ByVal sName, MAX_PATH, 0, 0)
        sName = GetStrFromBufferA(sName)
        isMalloc.Free ByVal lpsz

      ' The 1st UINT (Long) of the array points to the location
      ' (uOffset bytes) to the ANSII str in the pidl.
      Case STRRET_OFFSET:
        MoveMemory uOffset, lpStr.CStr(0), 4
        sName = GetStrFromPtrA(pidlRel + uOffset)
      
      ' The display name is returned in cStr.
      Case STRRET_CSTR:
        sName = GetStrFromPtrA(VarPtr(lpStr.CStr(0)))
    
    End Select
  End If
  
  GetFolderDisplayName = sName

Out:
End Function

' If successful, returns the relative pidl of the child namespace object from
' its SHGDN_INFOLDER name that resides in the specified parent folder.
' Returns 0 otherwise. (ISF.ParseDisplayName needs SHGDN_FORPARSING)

' Calling proc is responsible for freeing the returned pidl..

Public Function GetPIDLFromDisplayName(hWnd As Long, isfParent As IShellFolder, sName As String) As Long
  Dim grfFlags As Long
  Dim ieidl As IEnumIDList
  Dim pidlrelChild As Long
  
  grfFlags = SHCONTF_FOLDERS Or SHCONTF_NONFOLDERS Or SHCONTF_INCLUDEHIDDEN
  
  ' Creates an enumeration object for the parent folder's IShellFolder
  If (isfParent.EnumObjects(hWnd, grfFlags, ieidl) = NOERROR) Then
    
    ' Walk the contents of the enumeration object
    Do While (ieidl.Next(1, pidlrelChild, 0) = NOERROR)
      If (sName = GetFolderDisplayName(isfParent, pidlrelChild, SHGDN_INFOLDER)) Then
        GetPIDLFromDisplayName = pidlrelChild
        Exit Function
      End If
      isMalloc.Free ByVal pidlrelChild
    Loop
  
  End If   ' EnumObjects
  
End Function


