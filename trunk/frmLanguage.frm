VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "COMCTL32.OCX"
Begin VB.Form frmTranslations 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "Please select a translation..."
   ClientHeight    =   1440
   ClientLeft      =   45
   ClientTop       =   315
   ClientWidth     =   4080
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1440
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin ComctlLib.ListView lvw 
      Height          =   1455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   2566
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   2
      BeginProperty ColumnHeader(1) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Language"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(2) {0713E8C7-850A-101B-AFC0-4210102A8DA7} 
         SubItemIndex    =   1
         Key             =   ""
         Object.Tag             =   ""
         Text            =   "Translation author"
         Object.Width           =   2999
      EndProperty
   End
End
Attribute VB_Name = "frmTranslations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub lvw_DblClick()
 With Me.lvw
  If .SelectedItem Is Nothing Then Exit Sub
  If .SelectedItem.Selected = False Then Exit Sub
  Me.Hide
 End With
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
 If KeyAscii = vbKeyEscape Then Unload Me
End Sub
