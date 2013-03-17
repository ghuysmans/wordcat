Attribute VB_Name = "BrowseFolder"
'Browse for folder, by Alex Kail
'http://www.freevbcode.com/ShowCode.asp?ID=3064


Option Explicit

Private Type BROWSEINFO
 lngHwnd As Long
 pidlRoot As Long
 pszDisplayName As String
 lpszTitle As Long
 ulFlags As Long
 lpfnCallback As Long
 lParam As Long
 iImage As Long
End Type

Private Const BIF_RETURNONLYFSDIRS = 1
Private Const MAX_PATH = 260

Private Declare Sub CoTaskMemFree Lib "ole32.dll" (ByVal hMem As Long)
Private Declare Function lstrcat Lib "Kernel32" Alias "lstrcatA" (ByVal lpString1 As String, ByVal lpString2 As String) As Long
Private Declare Function SHBrowseForFolder Lib "shell32" (lpbi As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDList Lib "shell32" (ByVal pidList As Long, ByVal lpBuffer As String) As Long

Public Function BrowseForFolder(ByVal lngHwnd As Long, ByVal strPrompt As String) As String
 Dim intNull As Integer, strPath As String, udtBI As BROWSEINFO
 Dim lngIDList As Long, lngResult As Long
 
 On Error GoTo ehBrowseForFolder 'Trap for errors

 'Set API properties (housed in a UDT)
 With udtBI
  .lngHwnd = lngHwnd
  .lpszTitle = lstrcat(strPrompt, "")
  .ulFlags = BIF_RETURNONLYFSDIRS
 End With

 'Display the browse folder...
 lngIDList = SHBrowseForFolder(udtBI)

 If lngIDList <> 0 Then
  'Create string of nulls so it will fill in with the path
  strPath = String(MAX_PATH, 0)

  'Retrieves the path selected, places in the null
   'character filled string
  lngResult = SHGetPathFromIDList(lngIDList, strPath)

  'Frees memory
  Call CoTaskMemFree(lngIDList)

  'Find the first instance of a null character,
   'so we can get just the path
  intNull = InStr(strPath, vbNullChar)
  'Greater than 0 means the path exists...
  If intNull > 0 Then
   'Set the value
   strPath = Left(strPath, intNull - 1)
  End If
 End If

 'Return the path name
 BrowseForFolder = strPath
 Exit Function 'Abort

ehBrowseForFolder:

 'Return no value
 BrowseForFolder = Empty

End Function
