VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmMain 
   ClientHeight    =   8400
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9825
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   560
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   655
   StartUpPosition =   3  'Windows Default
   Begin ComctlLib.Toolbar tbr 
      Align           =   1  'Align Top
      Height          =   420
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   741
      ButtonWidth     =   635
      ButtonHeight    =   582
      AllowCustomize  =   0   'False
      Wrappable       =   0   'False
      Appearance      =   1
      ImageList       =   "iml"
      _Version        =   327682
      BeginProperty Buttons {0713E452-850A-101B-AFC0-4210102A8DA7} 
         NumButtons      =   15
         BeginProperty Button1 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "open"
            Object.ToolTipText     =   "Open another folder..."
            Object.Tag             =   ""
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button3 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "sort"
            Object.ToolTipText     =   "Sort items alphabetically"
            Object.Tag             =   ""
            ImageIndex      =   13
         EndProperty
         BeginProperty Button4 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "up"
            Object.ToolTipText     =   "Move the selected item up"
            Object.Tag             =   ""
            ImageIndex      =   8
         EndProperty
         BeginProperty Button5 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "down"
            Object.ToolTipText     =   "Move the selected item down"
            Object.Tag             =   ""
            ImageIndex      =   7
         EndProperty
         BeginProperty Button6 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "undo"
            Object.ToolTipText     =   "Undo last changes"
            Object.Tag             =   ""
            ImageIndex      =   14
         EndProperty
         BeginProperty Button7 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button8 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "check"
            Object.ToolTipText     =   "Check"
            Object.Tag             =   ""
            ImageIndex      =   4
         EndProperty
         BeginProperty Button9 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   "run"
            Object.ToolTipText     =   "Run!"
            Object.Tag             =   ""
            ImageIndex      =   9
         EndProperty
         BeginProperty Button10 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button11 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "tools"
            Object.ToolTipText     =   "Tools"
            Object.Tag             =   ""
            ImageIndex      =   6
         EndProperty
         BeginProperty Button12 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Enabled         =   0   'False
            Key             =   ""
            Object.Tag             =   ""
            Style           =   3
            MixedState      =   -1  'True
         EndProperty
         BeginProperty Button13 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "lang"
            Object.ToolTipText     =   "Select a language"
            Object.Tag             =   ""
            ImageIndex      =   15
         EndProperty
         BeginProperty Button14 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "help"
            Object.ToolTipText     =   "Open the help file"
            Object.Tag             =   ""
            ImageIndex      =   12
         EndProperty
         BeginProperty Button15 {0713F354-850A-101B-AFC0-4210102A8DA7} 
            Key             =   "about"
            Object.ToolTipText     =   "About this program..."
            Object.Tag             =   ""
            ImageIndex      =   11
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin VB.CommandButton cmdDebug 
      Caption         =   "DEBUG"
      Height          =   435
      Left            =   8460
      TabIndex        =   3
      Tag             =   "dnt"
      Top             =   480
      Visible         =   0   'False
      Width           =   1275
   End
   Begin ComctlLib.TreeView tvw 
      Height          =   4695
      Left            =   0
      TabIndex        =   0
      Top             =   420
      Width           =   3315
      _ExtentX        =   5847
      _ExtentY        =   8281
      _Version        =   327682
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   1
      ImageList       =   "iml"
      Appearance      =   0
   End
   Begin ComctlLib.ListView lvw 
      Height          =   4695
      Left            =   3360
      TabIndex        =   1
      Top             =   420
      Width           =   6135
      _ExtentX        =   10821
      _ExtentY        =   8281
      SortKey         =   3
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      SmallIcons      =   "iml"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   0
      NumItems        =   4
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Name"
         Object.Width           =   3704
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Alignment       =   1
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Size"
         Object.Width           =   1323
      EndProperty
      BeginProperty ColumnHeader(3) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   2
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Last Modified"
         Object.Width           =   3069
      EndProperty
      BeginProperty ColumnHeader(4) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   3
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Order"
         Object.Width           =   0
      EndProperty
   End
   Begin VB.ListBox lstLogs 
      Height          =   1620
      Left            =   0
      TabIndex        =   2
      Top             =   5160
      Width           =   9495
   End
   Begin ComctlLib.ImageList iml 
      Left            =   0
      Top             =   6540
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   327682
      BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
         NumListImages   =   15
         BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":058A
            Key             =   "open"
         EndProperty
         BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":08DC
            Key             =   "file"
         EndProperty
         BeginProperty ListImage3 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0C2E
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage4 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":0F80
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":12D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1624
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1976
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":1CC8
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":201A
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":236C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":26BE
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2A10
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":2D62
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":30B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
            Picture         =   "frmMain.frx":3406
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mnuTools 
      Caption         =   "Tools"
      Visible         =   0   'False
      Begin VB.Menu mnuT_rmHF 
         Caption         =   "Remove headers/footers"
      End
      Begin VB.Menu mnuT_ctF 
         Caption         =   "Count files"
      End
      Begin VB.Menu mnuT_Reset 
         Caption         =   "Reset order"
      End
      Begin VB.Menu mnuT_RemTmp 
         Caption         =   "Remove temp. files"
      End
      Begin VB.Menu mnuT_ClearLogs 
         Caption         =   "Clear logs"
      End
   End
   Begin VB.Menu mnuMerge 
      Caption         =   "Merge"
      Visible         =   0   'False
      Begin VB.Menu mnuM_fast 
         Caption         =   "fast method. No int. file"
      End
      Begin VB.Menu mnuM_slow 
         Caption         =   "slower method with int. files"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Inited As Boolean, LastNode As Node
Private MinWidth As Integer, MinHeight As Integer
Private WithEvents fe As clsEnumFiles, FE_Mode As EnumMode
Attribute fe.VB_VarHelpID = -1

Private Sub EnableRun(e As Boolean)
 Me.tbr.Buttons("run").Enabled = e
End Sub

Private Sub EnableControls(e As Boolean)
 Dim i As Integer
 With Me.tbr.Buttons
  For i = 1 To .Count - 2
   .Item(i).Enabled = e
  Next i
 End With
 Me.tvw.Enabled = e
 EnableRun False
End Sub

Private Sub Form_Resize()
 Dim r As RECT
 
 If Me.Width < MinWidth Then Me.Width = MinWidth
 If Me.Height < MinHeight Then Me.Height = MinHeight
 
 GetClientRect Me.hWnd, r
 Me.tvw.Height = r.b - Me.lstLogs.Height - Me.tbr.Height
 Me.tvw.Width = 0.4 * r.r
 Me.lvw.Height = Me.tvw.Height
 Me.lvw.Width = r.r - Me.tvw.Width - 8
 Me.lvw.Left = Me.tvw.Width + 4
 Me.lstLogs.Top = Me.tvw.Height + Me.tbr.Height
 Me.lstLogs.Width = r.r
End Sub

Private Sub Form_Load()
 Set fe = New clsEnumFiles
 Tr.TranslateForm Me
 Me.Caption = Prog.AppName & " " & App.LegalCopyright
 MinWidth = Me.Width: MinHeight = Me.Height
 ApplySize Me, MinHeight, MinWidth
 #If dbg = 1 Then
  Me.cmdDebug.Visible = True
  Me.lvw.ColumnHeaders(4).Width = 50
 #End If
End Sub


Private Sub LoadList(TargetDir As String)
 Me.lvw.ListItems.Clear
 FE_Mode = PopItems
 If fe.Enumerate(TargetDir, False) Then
  Addlog Tr.Translate("Something went wrong while enumerating files... Did you delete or rename anything?", True), vbCritical
  EndProgram
 End If
 ParseOrderFile TargetDir & Prog.OrderFilename, TargetDir, Me.lvw
End Sub

Public Sub InvokeTool(m As EnumMode, name As String, rec As Boolean)
 EnableControls False
 Addlog Tr.Translate("Invoked tool: $", True, name)
 Prog.Count_Files = 0
 Select Case m
  Case T_RmHF
   If Confirm("modify all documents? This action can't be undone!") = False Then
    Addlog Tr.Translate("Cancelled.", True)
    EnableControls True
    Exit Sub
   End If
  Case T_Reset
   If Confirm("reset order of all files? This action can't be undone!") = False Then
    Addlog Tr.Translate("Cancelled.", True)
    EnableControls True
    Exit Sub
   End If
 End Select
 FE_Mode = m
 If fe.Enumerate(Prog.Target, rec) Then
  Addlog Tr.Translate("Can't enumerate files into $!", True, Prog.Target), vbExclamation
  EnableControls True
  Exit Sub
 End If
 Select Case m
  Case T_CtF, T_RmTmp, T_Reset: Addlog Tr.Translate("Done. # file#{,s} processed.", True, Prog.Count_Files), vbInformation
  Case T_RmHF: Addlog Tr.Translate("Headers/footers removed. # file#{,s} processed.", True, Prog.Count_Files), vbInformation
 End Select
 EnableControls True
End Sub

Private Sub tvw_Collapse(ByVal Node As ComctlLib.Node): Node.Image = "closed": End Sub
Private Sub tvw_Expand(ByVal Node As ComctlLib.Node): Node.Image = "open": End Sub

Private Sub tvw_NodeClick(ByVal Node As ComctlLib.Node)
 If Inited = True Then SaveOrder fe, Me.lvw, LastNode.Key & Prog.OrderFilename
 Set LastNode = Node
 LoadList Node.Key
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
 If LastNode Is Nothing Then Exit Sub
 SaveOrder fe, Me.lvw, LastNode.Key & Prog.OrderFilename
 SaveSize Me
End Sub

Private Sub Form_Activate()
 Dim notfirst As Boolean
 If Inited Then Exit Sub

 Do
  If notfirst = True Then _
    Addlog Tr.Translate("This folder can't be accessed: $", True, Prog.Target), vbExclamation

  frmTarget.Show 1, Me
  If frmTarget.PressedOK = False Then EndProgram
  Unload frmTarget

  With Me.tvw.Nodes
   .Clear
   .Add , tvwChild, Prog.Target & "\", Tr.Translate("[Target]", True), "closed"
  End With

  notfirst = True
 Loop While fe.Enumerate(Prog.Target, True)

 ExpandAll Me.tvw
 On Error Resume Next
 Set Me.tvw.SelectedItem = Me.tvw.Nodes(1)
 tvw_NodeClick Me.tvw.Nodes(1)
 On Error GoTo 0

 Inited = True
 Addlog Tr.Translate("Ready. Target: $", True, Prog.Target)
End Sub

Private Sub fe_ObjectFound(IsDirectory As Boolean, name As String, sz As Double, ft As String)
 If FE_Mode < Checking Then
  PopulateView fe, FE_Mode, IIf(FE_Mode = PopTree, Me.tvw, Me.lvw), IsDirectory, name, sz, ft
 ElseIf FE_Mode = Checking Then
  CheckFiles_CB fe, IsDirectory, name
 ElseIf FE_Mode = OrderFiles Then
  MergeFilesS_CB fe, IsDirectory, name
 Else
  Tools_CB fe, FE_Mode, IsDirectory, name
 End If
End Sub

Private Sub lvw_DblClick()
 If SelectedItem(Me.lvw) Then ShellOpen Me.hWnd, Me.lvw.SelectedItem.Key
End Sub

Private Sub mnuM_slow_Click()
 SaveOrder fe, Me.lvw, LastNode.Key & Prog.OrderFilename
 EnableControls False
 FE_Mode = OrderFiles
 MergeFilesS Me.hWnd, fe, Prog.Target
 EnableControls True
End Sub

Private Sub mnuM_fast_Click()
 SaveOrder fe, Me.lvw, LastNode.Key & Prog.OrderFilename
 EnableControls False
 MergeFiles Me.hWnd, Prog.Target & "\", Prog.Tpl
 EnableControls True
End Sub

Private Sub mnuT_ctF_Click(): InvokeTool T_CtF, Me.mnuT_ctF.Caption, True: End Sub
Private Sub mnuT_rmHF_Click(): InvokeTool T_RmHF, Me.mnuT_rmHF.Caption, True: End Sub
Private Sub mnuT_Reset_Click(): InvokeTool T_Reset, Me.mnuT_Reset.Caption, True: End Sub
Private Sub mnuT_RemTmp_Click(): InvokeTool T_RmTmp, Me.mnuT_RemTmp.Caption, True: End Sub

Private Sub mnuT_ClearLogs_Click()
 Me.lstLogs.Clear
 Addlog Tr.Translate("Logs cleared.", True)
End Sub

Private Sub tbr_ButtonClick(ByVal Button As ComctlLib.Button)
 Dim e As Boolean
 Select Case Button.Key
  Case "open"
   SaveOrder fe, Me.lvw, LastNode.Key & Prog.OrderFilename
   Me.lvw.ListItems.Clear: Me.tvw.Nodes.Clear
   FE_Mode = PopTree
   Inited = False
   Form_Activate
  Case "check"
   SaveOrder fe, Me.lvw, LastNode.Key & Prog.OrderFilename
   EnableControls False
   FE_Mode = Checking
   e = CheckFiles(fe, Prog.Target)
   EnableControls True
   If e Then EnableRun True
  Case "undo"
   If SelectedItem(Me.tvw, False, False) = False Then Exit Sub
   If Confirm("discard your last changes?") = False Then Exit Sub
   EnableControls False
   LoadList Me.tvw.SelectedItem.Key
   EnableControls True
  Case "tools"
   SaveOrder fe, Me.lvw, LastNode.Key & Prog.OrderFilename
   PopupMenu mnuTools
  Case "up": OrderUp Me.lvw
  Case "down": OrderDn Me.lvw
  Case "sort": AlphaSort Me.lvw
  Case "run": PopupMenu mnuMerge
  Case "lang": ChooseLanguage
  Case "about"
   MsgBox Prog.AppName & " " & App.LegalCopyright & vbCrLf & vbCrLf & _
          Replace$(App.Comments, "  ", vbCrLf), vbInformation, Tr.Translate("About this program...", True)
  Case "help": ShellOpen Me.hWnd, "help" & Tr.CurrentTranslation & "\index.html"
 End Select
End Sub

Private Sub cmdDebug_Click()
 '''
End Sub
