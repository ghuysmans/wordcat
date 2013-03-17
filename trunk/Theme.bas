Attribute VB_Name = "Theme"
Option Explicit

Private Const MAX_PATH As Long = 260
Private Declare Function InitCommonControls Lib "comctl32.dll" () As Long
Private Declare Function GetModuleFileName Lib "kernel32.dll" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

Public Sub InitTheme()
 Dim fname As String, errStr As String, raw As String
 InitCommonControls
 fname = App.Path & "\" & App.EXEName & ".exe.manifest"
 If IsDebugging = False And Dir$(fname) = "" Then
  raw = "<?xml version='1.0' encoding='UTF-8' standalone='yes'?><assembly xmlns='urn:schemas-microsoft-com:asm.v1' manifestVersion='1.0'><assemblyIdentity version='" & _
        App.Major & "." & App.Minor & "." & App.Revision & ".0' processorArchitecture='X86' name='" & App.Title & "' type='win32' /><description>" & App.FileDescription & _
        "</description><dependency><dependentAssembly><assemblyIdentity type='win32' name='Microsoft.Windows.Common-Controls' version='6.0.0.0' processorArchitecture='X86' " & _
        "publicKeyToken='6595b64144ccf1df' language='*' /></dependentAssembly></dependency></assembly>"
  If WriteFile(fname, raw) Then
   Addlog "Can't create a manifest file: no theme will be applied.", vbExclamation
   Exit Sub
  End If
  Shell App.Path & "\" & App.EXEName & ".exe", vbNormalFocus
  EndProgram
 End If
End Sub

Public Function IsDebugging() As Boolean
 Dim buffer As String
 buffer = String$(MAX_PATH, 0)
 GetModuleFileName App.hInstance, buffer, MAX_PATH
 IsDebugging = (Right$(UCase$(Mid(buffer, 1, InStr(buffer, vbNullChar) - 1)), 8) = "\VB6.EXE")
End Function
