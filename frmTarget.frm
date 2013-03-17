VERSION 5.00
Begin VB.Form frmTarget 
   BorderStyle     =   4  'Fixed ToolWindow
   ClientHeight    =   780
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   780
   ScaleWidth      =   7455
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.ComboBox cmbTpl 
      Height          =   315
      Left            =   900
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   420
      Width           =   5655
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "OPEN"
      Enabled         =   0   'False
      Height          =   315
      Left            =   6600
      TabIndex        =   2
      Top             =   420
      Width           =   795
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "..."
      Height          =   315
      Left            =   6600
      TabIndex        =   0
      Tag             =   "dnt"
      Top             =   60
      Width           =   315
   End
   Begin VB.TextBox txtTarget 
      Height          =   315
      Left            =   60
      Locked          =   -1  'True
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   60
      Width           =   6495
   End
   Begin VB.Label lblTpl 
      Alignment       =   1  'Right Justify
      Caption         =   "Template: "
      Height          =   255
      Left            =   0
      TabIndex        =   4
      Top             =   480
      Width           =   855
   End
End
Attribute VB_Name = "frmTarget"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public PressedOK As Boolean
Private WithEvents fe As clsEnumFiles
Attribute fe.VB_VarHelpID = -1


Private Sub cmdBrowse_Click()
 Dim t As String
 t = BrowseFolder.BrowseForFolder(Me.hWnd, Me.Caption & ":")
 If CheckTarget(t) Then Me.txtTarget.Text = t _
 Else CheckTarget Me.txtTarget.Text
End Sub

Private Sub cmdOk_Click()
 Prog.Target = Me.txtTarget.Text
 If Me.cmbTpl.ListIndex = 0 Then Prog.Tpl = "" _
 Else Prog.Tpl = Me.cmbTpl.Text
 Prog.SaveSettings
 PressedOK = True
 Me.Hide
End Sub

Private Sub Form_Activate()
 Set fe = New clsEnumFiles
 Tr.TranslateForm Me
 Me.Caption = Tr.Translate("Target directory", True)
 Me.txtTarget.Text = Prog.Target
 CheckTarget Me.txtTarget.Text
End Sub


Private Sub CheckEsc(ascii As Integer)
 If ascii = vbKeyEscape Then Unload Me
End Sub

Private Sub txtTarget_KeyPress(KeyAscii As Integer): CheckEsc KeyAscii: End Sub
Private Sub cmdBrowse_KeyPress(KeyAscii As Integer): CheckEsc KeyAscii: End Sub
Private Sub cmbTpl_KeyPress(KeyAscii As Integer): CheckEsc KeyAscii: End Sub
Private Sub cmdOk_KeyPress(KeyAscii As Integer): CheckEsc KeyAscii: End Sub


Private Sub GetTemplates(t As String)
 With Me.cmbTpl
  .Clear
  .AddItem "[no template]"
  .ListIndex = 0
  fe.Enumerate t, True
  On Error Resume Next
  .ListIndex = 1 'if it fails --> 0
  On Error GoTo 0
  .Enabled = True
 End With
End Sub

Private Sub fe_ObjectFound(IsDirectory As Boolean, name As String, sz As Double, ft As String)
 If IsDirectory Then Exit Sub
 If fe.CheckExtension(name, "dot") = False Then Exit Sub
 Me.cmbTpl.AddItem name
End Sub

Private Function CheckTarget(t As String) As Boolean
 With Me.cmdOk
  If Len(t) Then
   .Enabled = True
   GetTemplates t
   CheckTarget = True
  Else
   Me.cmbTpl.Enabled = False
   .Enabled = False
  End If
 End With
End Function
