VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsEnumFiles"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'clsEnumFiles: VB6/VBA x64-compatible class to enumerate files via Windows APIs
'Copyright (C) 2013 Guillaume Huysmans
'
'This program is free software; you can redistribute it and/or
'modify it under the terms of the GNU Lesser General Public
'License as published by the Free Software Foundation; either
'version 3 of the License, or (at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE. See the GNU
'Lesser General Public License for more details.
'
'You should have received a copy of the GNU Lesser General Public
'License along with this program; if not, write to the Free Software
'Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA 02110-1301 USA


Option Explicit

Private Const MAX_PATH As Long = 260
Private Const INVALID_HANDLE_VALUE As Long = -1
Private Const FILE_ATTRIBUTE_DIRECTORY As Long = &H10


Private Type FILETIME
 dwLowDateTime As Long
 dwHighDateTime As Long
End Type

Private Type WIN32_FIND_DATA
 dwFileAttributes As Long
 ftCreationTime As FILETIME
 ftLastAccessTime As FILETIME
 ftLastWriteTime As FILETIME
 nFileSizeHigh As Long
 nFileSizeLow As Long
 dwReserved0 As Long
 dwReserved1 As Long
 cFileName As String * MAX_PATH
 cAlternate As String * 14
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


Private Declare Function FileTimeToLocalFileTime Lib "Kernel32" (lpFileTime As FILETIME, lpLocalFileTime As FILETIME) As Long
Private Declare Function FileTimeToSystemTime Lib "Kernel32" (lpFileTime As FILETIME, lpSystemTime As SYSTEMTIME) As Long
#If VBA7 Then
 Private Declare PtrSafe Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileA" (ByVal lpFileName As String, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
 Private Declare PtrSafe Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileA" (ByVal hFindFile As LongPtr, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
 Private Declare PtrSafe Function FindClose Lib "kernel32.dll" (ByVal hFindFile As LongPtr) As Long
#Else
 Private Declare Function FindFirstFile Lib "kernel32.dll" Alias "FindFirstFileA" (ByVal lpFileName As String, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
 Private Declare Function FindNextFile Lib "kernel32.dll" Alias "FindNextFileA" (ByVal hFindFile As Long, ByRef lpFindFileData As WIN32_FIND_DATA) As Long
 Private Declare Function FindClose Lib "kernel32.dll" (ByVal hFindFile As Long) As Long
#End If

Public Event ObjectFound(IsDirectory As Boolean, name As String, sz As Double, ft As String)
Public Halt As Boolean, DF As Boolean


Private Function FiletimeToString(ft As FILETIME) As String
 Dim lft As FILETIME, st As SYSTEMTIME
 'It's stored as an UTC time
 FileTimeToLocalFileTime ft, lft
 'Convert it to a usable structure
 FileTimeToSystemTime lft, st
 'Use VB's date functions to format it using the user's preferences
 FiletimeToString = CStr(DateSerial(st.wYear, st.wMonth, st.wDay)) & " " & _
                    CStr(TimeSerial(st.wHour, st.wMinute, st.wSecond))
End Function

Private Function GoEnum(Path As String, Recursive As Boolean) As Boolean
'Returns TRUE if anything went wrong, otherwise FALSE.
 Dim fd As WIN32_FIND_DATA, IsDir As Boolean, ft As String
 Dim name As String, Res As Boolean
 #If VBA7 Then
  Dim hFind As LongPtr
 #Else
  Dim hFind As Long
 #End If
 hFind = FindFirstFile(Path & "*.*", fd)
 If hFind <> INVALID_HANDLE_VALUE Then
  Do
   name = Left$(fd.cFileName, InStr(1, fd.cFileName, vbNullChar) - 1)
   If Len(name) > 0 And (name <> ".") And (name <> "..") Then
    ft = FiletimeToString(fd.ftLastWriteTime)
    If (fd.dwFileAttributes And FILE_ATTRIBUTE_DIRECTORY) = 0 Then
     'file
     RaiseEvent ObjectFound(False, Path & name, fd.nFileSizeHigh * 2 ^ 32 + fd.nFileSizeLow, ft)
    Else
     'directory
     name = name & "\"
     If DF Then
      RaiseEvent ObjectFound(True, Path & name, 0, ft)
      If Recursive And (Halt = False) Then Res = Res Or (GoEnum(Path & name, True))
     Else
      If Recursive And (Halt = False) Then Res = Res Or (GoEnum(Path & name, True))
      RaiseEvent ObjectFound(True, Path & name, 0, ft)
     End If
    End If
   End If
  Loop While (FindNextFile(hFind, fd) <> 0) And (Halt = False) And (Res = False)
  FindClose hFind
  GoEnum = Res Or Halt
 Else
  GoEnum = True
 End If
End Function


Public Function Enumerate(ByVal Path As String, Optional Recursive As Boolean = False, Optional DirFirst As Boolean = True) As Boolean
 Path = Replace$(Path, "/", "\")
 If Right$(Path, 1) <> "\" Then Path = Path & "\"
 Halt = False: DF = DirFirst
 Enumerate = GoEnum(Path, Recursive)
End Function

'The API doesn't handle any other filter than *.*, so I have to check the extension myself. Thanks, MS!
Public Function CheckExtension(Filename As String, Extension As String) As Boolean
'Returns TRUE if it matches, FALSE if it doesn't.
 Dim s As String
 On Error GoTo err
 s = Right$(Filename, Len(Extension) + 1)
 If LCase$(s) = "." & LCase$(Extension) Then CheckExtension = True
 'If an error has occured (too short name) OR if it doesn't match, default=FALSE.
err:
End Function

Public Function GetCanonicalFolderName(ByVal fname As String) As String
 fname = Replace$(fname, "/", "\")
 If Right$(fname, 1) = "\" Then
  GetCanonicalFolderName = Left$(fname, Len(fname) - 1)
 Else
  GetCanonicalFolderName = fname
 End If
End Function
