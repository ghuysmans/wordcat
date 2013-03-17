Attribute VB_Name = "Prog"
Option Explicit

Public Const OrderFilename As String = "order.dat"
Public Const IntermFilename As String = "order.tmp.doc"
Public Const DestFilename As String = "merge.doc"
Private Const SW_SHOW As Long = 5
Private Const wdStory As Integer = 6

Public Enum EnumMode
 PopTree
 PopItems
 Checking
 T_RmHF
 T_CtF
 OrderFiles
 T_RmTmp
 T_Reset
End Enum

Public Type RECT
 l As Long
 t As Long
 r As Long
 b As Long
End Type

Public Declare Function GetClientRect Lib "user32.dll" (ByVal hWnd As Long, ByRef r As RECT) As Boolean
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function MoveFile Lib "kernel32.dll" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long

Public Target As String, Tpl As String
Public Count_Files As Integer
Public Tr As clsTranslate
Private Check_Named As Collection
Private Check_All As Collection
Private WordObj As Object


Public Function Confirm(s As String) As Boolean
 Confirm = MsgBox(Tr.Translate("Are you sure to " & s, True), vbQuestion Or vbYesNo, Tr.Translate("Confirmation", True)) = vbYes
End Function

Public Sub ShellOpen(hWnd As Long, Target As String)
 ShellExecute hWnd, "open", Target, "", "", SW_SHOW
End Sub

Public Sub Addlog(s As String, Optional Style As VbMsgBoxStyle = -1)
 With frmMain.lstLogs
  .AddItem Now & "   " & s
  On Error Resume Next
  .ListIndex = .ListCount - 1
  On Error GoTo 0
 End With
 If Style <> -1 Then MsgBox s, Style
End Sub

Public Function AppName() As String
 AppName = App.Title & " v" & App.Major & "." & App.Minor & IIf(App.Revision, " r" & App.Revision, "")
End Function

Public Function FileExists(fname As String) As Boolean
 On Error Resume Next
 FileExists = CBool(GetAttr(fname))
End Function

Private Sub SafeKill(fname As String)
 On Error Resume Next
 Kill fname
End Sub

Private Function ReadFile(fname As String) As String
 Dim ha As Integer: ha = FreeFile
 If FileExists(fname) = False Then Exit Function
 Open fname For Binary As #ha
 ReadFile = Space$(LOF(ha))
 Get #ha, , ReadFile
 Close #ha
End Function

Public Function WriteFile(fname As String, Contents As String) As Boolean
 Dim ha As Integer: ha = FreeFile
 SafeKill fname
 Open fname For Binary As #ha
 Put #ha, , Contents
 Close #ha
End Function


Public Function InterestingObject(fe As clsEnumFiles, IsDir As Boolean, short As String) As Boolean
 InterestingObject = (Left$(short, 1) <> "-") And (short <> IntermFilename) And (short <> DestFilename)
 If IsDir Or (InterestingObject = False) Then Exit Function
 InterestingObject = fe.CheckExtension(short, "doc")
End Function

Public Function SelectedItem(obj As Object, Optional m As Boolean = False, Optional fc As Boolean) As Boolean
 If obj.SelectedItem Is Nothing Then GoTo tell
 If fc Then SelectedItem = obj.SelectedItem.Selected Else SelectedItem = True
 If SelectedItem Then Exit Function
tell:
 MsgBox Tr.Translate("Please first select an item!", True), vbExclamation
End Function

Public Sub PopulateView(fe As clsEnumFiles, FE_Mode As EnumMode, obj As Object, IsDirectory As Boolean, name As String, sz As Double, ft As String)
 Dim parent As String, short As String, p As Integer, li As ListItem
 If IsDirectory And (FE_Mode = PopItems) Then name = Left$(name, Len(name) - 1)
 p = InStrRev(name, "\", Len(name) - 1)
 short = Mid$(name, p + 1)
 If FE_Mode = PopTree Then
  If IsDirectory = False Then Exit Sub
  parent = Mid$(name, 1, p) 'not really useful but easier to read
  short = Left$(short, Len(short) - 1)
  obj.Nodes.Add parent, tvwChild, name, short, "closed"
 Else
  If IsDirectory Then
   Set li = obj.ListItems.Add(, name, short, , "closed")
   li.SubItems(1) = "[DIR]"
   li.SubItems(2) = ft
  Else
   If InterestingObject(fe, False, name) Then
    Set li = obj.ListItems.Add(, name, short, , "file")
    li.SubItems(1) = FormatNumber(sz, 0, vbFalse, vbFalse, vbTrue)
    li.SubItems(2) = ft
   End If
  End If
 End If
End Sub

Public Sub ExpandAll(tree As TreeView)
 Dim n As Node
 For Each n In tree.Nodes
  n.Expanded = True
 Next n
End Sub

Private Function FormatOrder(o As Integer) As String
 FormatOrder = Right$("0000" & CStr(o), 5)
End Function

Public Sub ParseOrderFile(Filename As String, BaseDir As String, lvw As ListView)
 Dim arr() As String, i As Integer, b As Integer, sort As String, li As ListItem
 arr = Split(ReadFile(Filename), vbCrLf): b = UBound(arr)
 On Error Resume Next 'for non-existing files
 For i = 0 To b
  sort = FormatOrder(i)
  lvw.ListItems(BaseDir & arr(i)).SubItems(3) = sort
 Next i
 lvw.SortKey = 3
 lvw.Sorted = True
End Sub

Public Sub SaveOrder(fe As clsEnumFiles, lvw As ListView, dest As String)
 Dim raw As String: raw = CreateOrderFile(fe, lvw)
 If raw = ReadFile(dest) Then Exit Sub 'nothing to do
 If WriteFile(dest, raw) Then
  Addlog Tr.Translate("Can't write the objects order into $", True, dest), vbCritical
  EndProgram
 End If
End Sub

Public Function CreateOrderFile(fe As clsEnumFiles, lvw As ListView) As String
 Dim raw As String, li As ListItem
 For Each li In lvw.ListItems
  If InterestingObject(fe, IIf(li.SmallIcon = "closed", True, False), li.Text) Then _
    raw = raw & li.Text & vbCrLf
 Next li
 If Len(raw) = 0 Then Exit Function
 CreateOrderFile = Left$(raw, Len(raw) - 2)
End Function

Public Sub AlphaSort(lvw As ListView)
 lvw.SortKey = 0
 lvw.Sorted = True
End Sub

Public Function CheckFiles(fe As clsEnumFiles, Target As String)
 Dim i As Integer, x As String, Check_Problems As Collection
 Set Check_All = New Collection
 Set Check_Named = New Collection
 fe.Enumerate Target, True
 Set Check_Problems = New Collection
 On Error Resume Next
 For i = 1 To Check_All.Count
  x = Check_Named(Check_All(i))
  If err Then
   Check_Problems.Add Check_All(i)
   err.Clear
  End If
 Next i
 x = ""
 For i = 1 To Check_Problems.Count
  x = x & Check_Problems(i) & vbCrLf
  Addlog Check_Problems(i)
 Next i
 If Len(x) Then
  MsgBox Tr.Translate("Some files/directories are not present in lists: to avoid skipping them," & vbCrLf & _
            "please browse to their parent directory. These files/directories are:" & vbCrLf & _
            vbCrLf & "$", True, x), vbExclamation, Tr.Translate("Check", True)
 Else
  MsgBox Tr.Translate("Everything seems to be fine!", True), vbInformation, Tr.Translate("Check", True)
  CheckFiles = True
 End If
End Function

Public Sub CheckFiles_CB(fe As clsEnumFiles, IsDirectory As Boolean, name As String)
 Dim p As Integer, short As String, parent As String
 Dim arr() As String, i As Integer, b As Integer
 If IsDirectory Then name = Left$(name, Len(name) - 1)
 p = InStrRev(name, "\")
 short = Mid$(name, p + 1)
 If InterestingObject(fe, True, short) = False Then Exit Sub
 If IsDirectory Then
  Check_All.Add name, name
 Else
  parent = Mid$(name, 1, p)
  If InterestingObject(fe, False, short) Then
   Check_All.Add name, name
  ElseIf short = OrderFilename Then
   arr = Split(ReadFile(name), vbCrLf)
   b = UBound(arr)
   For i = 0 To b
    Check_Named.Add parent & arr(i), parent & arr(i)
   Next i
  End If
 End If
End Sub

Public Sub MergeFilesS(hWnd As Long, fe As clsEnumFiles, Target As String)
 Count_Files = 0
 If fe.Enumerate(Target, True, False) Then
  Addlog Tr.Translate("This folder can't be accessed: $", True, Prog.Target), vbExclamation
  Exit Sub
 End If
 MergeFilesS_CB fe, True, Target & "\" 'forced 1st level
 GetWordObject
 If fe.Halt = False Then
  MoveFile Target & "\" & IntermFilename, Target & "\" & DestFilename 'rename
  Addlog Tr.Translate("Done. # file#{,s} processed.", True, Count_Files), vbInformation
  'If Confirm("keep intermediary files? They won't be deleted automatically.") = True Then _
  '   frmMain.InvokeTool T_RmTmp, frmMain.mnuT_RemTmp.Caption, True
  ShellOpen hWnd, Target & "\" & DestFilename
 End If
End Sub

Public Sub MergeFilesS_CB(fe As clsEnumFiles, IsDirectory As Boolean, name As String)
 Dim p As Integer, parent As String, src As String
 Dim arr() As String, u As Integer, i As Integer, d As Object
 If IsDirectory = False Then Exit Sub
 arr = Split(ReadFile(name & OrderFilename), vbCrLf)
 u = UBound(arr)
 If u = -1 Then Exit Sub 'don't save anything
 On Error Resume Next
 Set d = WordObj.Documents.Add(Prog.Tpl)
 If err Then
  Addlog Tr.Translate("Can't use the template $", True, Prog.Tpl), vbExclamation
  fe.Halt = True
  Exit Sub
 End If
 On Error GoTo 0
 WordObj.Selection.EndKey Unit:=wdStory
 For i = 0 To u
  src = name & arr(i)
  'If it is a directory, we have to use the file named IntermFilename into it.
  If (GetAttr(src) And vbDirectory) = vbDirectory Then
   If FileExists(src & "\" & IntermFilename) = False Then
    Addlog Tr.Translate("Empty folder: $", True, src)
    src = ""
   Else
    src = src & "\" & IntermFilename
   End If
  Else
   src = name & arr(i)
   Count_Files = Count_Files + 1
  End If
  If Len(src) Then
   On Error Resume Next
   WordObj.Selection.InsertFile src, , , False, False
   If err Then
    Addlog Tr.Translate("Can't insert $", True, src), vbExclamation
    fe.Halt = True
    Exit Sub
   End If
   On Error GoTo 0
   WordObj.Selection.EndKey Unit:=wdStory
   Addlog Tr.Translate("Merged: $", True, name & arr(i))
  End If
 Next i
 On Error Resume Next
 d.SaveAs name & IntermFilename
 If err Then
  Addlog Tr.Translate("Can't save $", True, name & IntermFilename), vbExclamation
  fe.Halt = True
  Exit Sub
 End If
 On Error GoTo 0
 Addlog Tr.Translate("Saved: $", True, name & IntermFilename)
 d.Close: Set d = Nothing
End Sub

Private Function MergeFiles_FastRec(d As Object, Target As String) As Boolean
 Dim arr() As String, u As Integer, i As Integer, src As String
 arr = Split(ReadFile(Target & OrderFilename), vbCrLf): u = UBound(arr)
 For i = 0 To u
  src = Target & arr(i)
  If (GetAttr(src) And vbDirectory) = vbDirectory Then
   If MergeFiles_FastRec(d, src & "\") Then
    MergeFiles_FastRec = True
    Exit Function
   End If
  Else
   On Error Resume Next
   WordObj.Selection.InsertFile src, , , False, False
   WordObj.Selection.EndKey Unit:=wdStory
   If err Then
    Addlog Tr.Translate("Can't insert $", True, src), vbExclamation
    MergeFiles_FastRec = True
    Exit Function
   End If
   On Error GoTo 0
   Count_Files = Count_Files + 1
  End If
  Addlog Tr.Translate("Merged: $", True, src)
 Next i
End Function

Public Sub MergeFiles(hWnd As Long, Target As String, Tpl As String)
 Dim d As Object, dest As String
 On Error Resume Next
 Set d = WordObj.Documents.Add(Tpl)
 If err Then
  Addlog Tr.Translate("Can't use the template $", True, Prog.Tpl), vbExclamation
  Exit Sub
 End If
 On Error GoTo 0
 Count_Files = 0
 If MergeFiles_FastRec(d, Target) Then Exit Sub
 dest = Target & DestFilename
 d.SaveAs dest
 d.Close: Set d = Nothing
 GetWordObject
 Addlog Tr.Translate("Saved: $", True, dest)
 Addlog Tr.Translate("Done. # file#{,s} processed.", True, Count_Files)
 ShellOpen hWnd, dest
End Sub

Public Sub Tools_CB(fe As clsEnumFiles, FE_Mode As EnumMode, IsDirectory As Boolean, name As String)
 Dim d As Object, s As Object, hf As Object, short As String

 If IsDirectory Then Exit Sub
 short = Mid$(name, InStrRev(name, "\") + 1)
 If FE_Mode = T_Reset Then
  If short <> OrderFilename Then Exit Sub
 ElseIf FE_Mode = T_RmTmp Then
  If short <> IntermFilename Then Exit Sub
 Else
  If InterestingObject(fe, False, short) = False Then Exit Sub
 End If

 If (FE_Mode = T_CtF) Or (FE_Mode = T_Reset) Or (FE_Mode = T_RmHF) Then Count_Files = Count_Files + 1

 Select Case FE_Mode
  Case T_Reset: SafeKill name
  Case T_RmTmp
   If short = IntermFilename Then
    SafeKill name
    Count_Files = Count_Files + 1
   End If
  Case T_RmHF
   Set d = WordObj.Documents.Open(name)
   For Each s In d.Sections
    For Each hf In s.Headers
     hf.Range.Delete
    Next hf
    For Each hf In s.Footers
     hf.Range.Delete
    Next hf
   Next s
   d.Save: d.Close
   Set d = Nothing
   Addlog Tr.Translate("Cleaned: $", True, name)
 End Select
End Sub


Private Sub Xchg(lvw As ListView, a As Integer, b As Integer)
'only works with a<b
 Dim tmp As ListItem, i As Integer, k As String
 Set tmp = lvw.ListItems.Add(b + 1)
 With lvw.ListItems(a)
  tmp.Icon = .Icon
  k = .Key
  tmp.SmallIcon = .SmallIcon
  tmp.Tag = .Tag
  tmp.Text = .Text
  For i = 1 To lvw.ColumnHeaders.Count - 1
   tmp.SubItems(i) = .SubItems(i)
  Next i
 End With
 lvw.ListItems.Remove a
 tmp.Key = k
End Sub

Public Sub OrderUp(lvw As ListView)
 Dim s As Integer, li As ListItem
 If SelectedItem(lvw) = False Then Exit Sub
 s = lvw.SelectedItem.Index
 lvw.Sorted = False
 With lvw.ListItems
  If .Count = 1 Or s = 1 Then Exit Sub
  If InterestingObject(Nothing, True, .Item(s - 1).Text) = False Then Exit Sub
  Xchg lvw, s - 1, s
  .Item(s - 1).Selected = True
 End With
 lvw.SetFocus
End Sub

Public Sub OrderDn(lvw As ListView)
 Dim s As Integer
 If SelectedItem(lvw) = False Then Exit Sub
 s = lvw.SelectedItem.Index
 lvw.Sorted = False
 With lvw.ListItems
  If .Count = 1 Or s = .Count Then Exit Sub
  Xchg lvw, s, s + 1
  .Item(s + 1).Selected = True
 End With
 lvw.SetFocus
End Sub


Public Sub SaveSettings()
 SaveSetting App.EXEName, "Settings", "TargetDir", Target
 SaveSetting App.EXEName, "Settings", "Template", Tpl
End Sub

Private Sub LoadSettings()
 Target = GetSetting(App.EXEName, "Settings", "TargetDir", "")
 Tpl = GetSetting(App.EXEName, "Settings", "Template", "")
 Tr.LoadTranslation GetSetting(App.EXEName, "Settings", "Translation", "\en.trn")
End Sub

Public Sub SaveSize(frm As Form)
 SaveSetting App.EXEName, frm.name, "H", frm.Height
 SaveSetting App.EXEName, frm.name, "W", frm.Width
End Sub

Public Sub ApplySize(frm As Form, dH As Integer, dW As Integer)
 frm.Height = GetSetting(App.EXEName, frm.name, "H", dH)
 frm.Width = GetSetting(App.EXEName, frm.name, "W", dW)
End Sub

Public Sub ChooseLanguage()
 Dim t As String: t = Tr.AskTranslation(frmMain)
 If Len(t) Then
  SaveSetting App.EXEName, "Settings", "Translation", t
  Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus
  EndProgram
 End If
End Sub

Public Sub GetWordObject()
 Set WordObj = Nothing 'we can't reuse the object after a merging process
 On Error Resume Next
 Set WordObj = CreateObject("Word.Application")
 If err Then
  MsgBox Tr.Translate("The Microsoft Word library can't be loaded. Is it installed?", True), vbCritical
  End
 End If
 On Error GoTo 0
End Sub

Private Sub Main()
 Theme.InitTheme

 Set Tr = New clsTranslate
 Tr.LoadModel "\en.trn"
 LoadSettings

 GetWordObject
 frmMain.Show
End Sub

Public Sub EndProgram()
 Set WordObj = Nothing
 Set Tr = Nothing
 End
End Sub
