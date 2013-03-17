VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTranslate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'clsTranslate, a translation class supporting almost every language
'Copyright (C) 2013 Guillaume HUYSMANS
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
'
'
'Warning: if you're trying to port this lib to ASP, don't forget
'that InStr doesn't return the same value when nothing is found!


Option Explicit

Private Type TCond
 min As Integer
 max As Integer
 md As Integer
 def As Boolean
End Type

Private Type TLangFile
 FileName As String
 Program As String
 Language As String
 Author As String
 GrammNumbers() As TCond
 Labels() As String
 LabelsCt As Long
 Loaded As Boolean
End Type

Private Enum EErrCode
 'ParseTranslation
 errMissingHeaderFields
 errIncompatible
 errInvGnList
 errNoLabel
 'AnalyzeGN
 errNoCondition
 errInvCondition
 'Fmt
 errNonExParam
 errParamNotNum
 errMissingClBrace
 errUnknownGn
 'Translate
 errNoTranslation
 errNoLoadedModel
 errNoLoadedTranslation
 'LoadTranslation
 errCannotRead
 errDiffLabelsCount
End Enum

Private p_model As TLangFile, p_transl As TLangFile
Private WithEvents fe As clsEnumFiles
Attribute fe.VB_VarHelpID = -1


'################ Public variables

Public Passthrough As Boolean, AutoInsert As Boolean, TrVersion As Integer
Public Warn_NotString As Boolean
Public ThousandsSep As Boolean, DigitsAfterDecimal As Integer, NoTrailingZero As Boolean
Public TargetFolder As String, TargetExtension As String, IgnoreTag As String


'################ Internally used function

Private Sub ThrowError(code As EErrCode, src As String, Optional details As String = "")
 Dim e() As String, s As String
 e = Split("missing header fields;incompatible translation file;invalid grammatical numbers list;there is no no label into it;" & _
            "no condition;invalid condition;" & _
            "reference to a non-existing parameter;this parameter is not a number;missing closing brace;unknown grammatical number;" & _
            "no translation found;no loaded model;no loaded translation;" & _
            "unreadable/inexistant translation file;labels counts are different between the model and the translation", ";")
 s = e(code) & IIf(Len(details), vbCrLf & "Details: " & details, "")
 Select Case code
  Case Is < errNoCondition: s = "Can't load the translation file: " & s
  Case Is < errNonExParam: s = "Can't parse the grammatical numbers definition: " & s
  Case Is < errNoTranslation: s = "Can't format the message: " & s
  Case Is < errCannotRead: s = "Can't translate the message: " & s
  Case Else: s = "Can't load the translation file: " & s
 End Select
 err.Raise vbObjectError Or code, "clsTranslate." & src, s & "."
End Sub

Private Function ParseTranslation(raw_s As String, Transl_Ver As Integer, ByRef lf As TLangFile, Optional HeadersOnly As Boolean = False) As Boolean
'no leading zeros for versions in the translation file! separator: comma.
'filename must be set after calling this function
 Dim llf As TLangFile, e As EErrCode 'local langfile, error code
 Dim raw() As String, rb As Long, i As Long 'raw string, raw lines, upper bound, iterator
 Dim curline As String, ml As Boolean, pml As Boolean 'current line, multiline, previous ml
 raw = Split(raw_s, vbCrLf): rb = UBound(raw)
 e = errMissingHeaderFields
 If rb < 5 Then GoTo er
 With llf
  e = errIncompatible
  If InStr(1, "," & raw(1) & ",", "," & CStr(Transl_Ver) & ",") = 0 Then GoTo er
  .Program = raw(0)
  .Language = raw(2)
  .Author = raw(3)
  e = errInvGnList
  If AnalyzeGN(raw(4), llf) Then GoTo er
  ReDim .Labels(rb - 5) 'maximum size: total - headers
 End With
 If HeadersOnly Then
  lf = llf
  Exit Function
 End If
 For i = 5 To rb
  curline = raw(i)
  If Len(curline) Then
   ml = (Right$(curline, 1) = "\") 'is multiline?
   If ml Then curline = Left$(curline, Len(curline) - 1)
   llf.Labels(llf.LabelsCt) = llf.Labels(llf.LabelsCt) & IIf(pml, vbCrLf, "") & curline
   If ml = False Then llf.LabelsCt = llf.LabelsCt + 1
   pml = ml 'save the multiline flag
  Else
   pml = False
  End If
 Next i
 e = errNoLabel
 If llf.LabelsCt = 0 Then GoTo er
 ReDim Preserve llf.Labels(llf.LabelsCt - 1)
 lf = llf
 Exit Function
er:
 ParseTranslation = True
 ThrowError e, "ParseTranslation"
End Function

Private Function AnalyzeGN(t As String, ByRef lf As TLangFile) As Boolean
'Returns TRUE if an error has occured, otherwise FALSE.
 Dim c() As String, n() As String, u As Integer
 Dim i As Integer, j As Integer, e As EErrCode

 If Len(t) = 0 Then
  e = errNoCondition
  GoTo er
 End If
 c = Split(t, ";"): u = UBound(c)
 ReDim lf.GrammNumbers(u)
 
 For i = 0 To u
  If c(i) = "d" Then
   lf.GrammNumbers(i).def = True
  Else
   n = Split(c(i), ",")
   If UBound(n) <> 2 Then
    e = errInvCondition
    GoTo er
   End If
   With lf.GrammNumbers(i)
    .min = -1: .max = -1: .md = 0
    If IsNumeric(n(0)) Then .min = CInt(n(0))
    If IsNumeric(n(1)) Then .max = CInt(n(1))
    If IsNumeric(n(2)) Then .md = CInt(n(2))
   End With
  End If
 Next i
 
 Exit Function
er:
 AnalyzeGN = True
 ThrowError e, "AnalyzeGN", t
End Function

Private Function GetNumberGN(ByVal n As Integer, ByRef lf As TLangFile) As Integer
'Returns -1 if nothing matches, else the category's ID.
 Dim i As Integer, t As Integer, u As Integer
 n = Abs(n): u = UBound(lf.GrammNumbers)
 For i = 0 To u
  With lf.GrammNumbers(i)
   If .def = False Then
    If .md Then t = n Mod .md Else t = n
    If (.min <> -1) And (t < .min) Then GoTo bad
    If (.max <> -1) And (t > .max) Then GoTo bad
   End If
   GetNumberGN = i
   Exit Function
bad:
  End With
 Next i
 
 GetNumberGN = -1
End Function

Private Function Fmt(lf As TLangFile, ByVal f As String, ByVal thsep As Boolean, ByVal warnint As Boolean, ByVal dad As Integer, ByVal ntz As Boolean, var As Variant) As String
'(The most complicated function in this library...)
'This function is NOT recursive.

    Dim param_ptr As Integer, i As Integer, param_ptr_copy As Integer  'param pointer (current), index, counter
    Dim p As Integer, p2 As Integer, p3 As Integer, p4 As Integer, l As Integer 'positions (for f manipulation), f's length
    Dim continue As Boolean, cc As String, id As String 'continue flag, current char, id (from the format string)
    Dim e As EErrCode, d As String, curp, variants() As String 'error code, error details, current parameter, variants (according to gramm numbers)
    Dim isstr As Boolean

    p = InStr(1, f, "//") 'comment start
    l = Len(f)
    f = Left$(f, IIf(p, p - 1, Len(f))) 'strip comments

    Do
        'Loop until we find an unescaped dollar/hash char
        '(sharp=diesis --> music, not IT)
        continue = True
        Do
            'Find the nearest interesting char
            p = InStr(1, f, "$")
            p2 = InStr(1, f, "#")
            isstr = CBool(p) 'and save somewhere its type
            If p = 0 Then p = p2

            If p = 1 Then
                continue = False
            Else
                If p = 0 Then
                    continue = False
                ElseIf Mid$(f, p - 1, 1) <> "\" Then
                    continue = False
                End If 'p = 0
            End If 'p = 1
        Loop While continue
        'Found?
        If p Then
            'Loop until we find something else than a digit (this will give us the requested ID)
            p2 = p + 1
            id = "": continue = True
            Do
                If p2 > l Then
                    continue = False
                Else
                    cc = Mid$(f, p2, 1)
                    If IsNumeric(cc) Then
                        id = id & cc
                    Else
                        continue = False
                    End If 'IsNumeric(cc)
                End If 'p2 > l
                If continue Then p2 = p2 + 1
            Loop While continue
            'p2-1 is the last char used to represent the id. This will be used in the next IF block
            If Len(id) Then param_ptr_copy = CInt(id) - 1 Else param_ptr_copy = param_ptr 'use the ID we've just read
            'Conditional block, we have to select the last one without erasing the pointer (before reading it)
            If Mid$(f, p2, 1) = "{" And param_ptr_copy Then param_ptr_copy = param_ptr_copy - 1
            
            'Load the parameter
            e = errNonExParam
            d = "parameter #" & param_ptr_copy + 1
            On Error GoTo er
            curp = var(param_ptr_copy)
            On Error GoTo 0
            
            If isstr Then
                param_ptr = param_ptr_copy
                f = Mid$(f, 1, p - 1) & curp & Mid$(f, p2)
                If warnint Then
                    If IsNumeric(curp) Then MsgBox "Warning: param #" & param_ptr + 1 & " is NOT a string!", vbExclamation
                End If
                l = l + Len(curp)
                param_ptr = param_ptr + 1 'increment the param pointer
            Else
                e = errParamNotNum
                d = "parameter #" & param_ptr_copy + 1 & " = """ & CStr(curp) & """"
                If IsNumeric(curp) = False Then GoTo er
                If Mid$(f, p2, 1) = "{" Then
                    'Find the closing brace
                    p4 = p2 + 1: p3 = 0: continue = True
                    Do
                        p4 = InStr(p4, f, "}")
                        If p4 Then
                            If Mid$(f, p4 - 1, 1) <> "\" Then continue = False
                            p3 = p4 'save it
                        Else
                            continue = False
                        End If 'p4
                    Loop While continue
                    'Error if missing
                    e = errMissingClBrace: d = ""
                    If p3 = 0 Then GoTo er
                    'Choose the right variant
                    i = GetNumberGN(CInt(curp), lf)
                    e = errUnknownGn: d = Abs(CInt(curp))
                    If i = -1 Then GoTo er
                    curp = Mid$(f, p2 + 1, p3 - p2 - 1)
                    'Ignore escaped commas - useless...
                    'curp = Replace$(curp, ",", "!,")
                    'curp = Replace$(curp, "\!", "")
                    'variants = Split(curp, "!,")
                    variants = Split(curp, ",")
                    'insert it
                    f = Mid$(f, 1, p - 1) & variants(i) & Mid$(f, p3 + 1)
                    'recompute the length
                    l = Len(f)
                Else
                    param_ptr = param_ptr_copy
                    'format the number - i represents ONLY HERE the number of DAD
                    i = dad
                    If ntz And (curp = CInt(curp)) Then i = 0 'no trailing zeros
                    curp = FormatNumber$(curp, i, , , IIf(thsep, vbTrue, vbFalse))
                    'insert it
                    f = Mid$(f, 1, p - 1) & curp & Mid$(f, p2)
                    'increment the param pointer
                    param_ptr = param_ptr + 1
                    'increment the length
                    l = l + Len(curp)
                End If 'Mid$(f, p2, 1) = "{"
            End If 'isstr
        End If 'p
    Loop While p

    Fmt = f
    Exit Function
er:
    ThrowError e, "Fmt", d
End Function

Private Function FindLabel(msg As String, exact As Boolean, lf As TLangFile) As Long
 Dim i As Long, l As Integer
 l = Len(msg)
 For i = 0 To lf.LabelsCt - 1
  'A way to avoid doing both tests on EACH line...
  If exact Then
   If lf.Labels(i) <> msg Then GoTo nx
  Else
   If Left$(lf.Labels(i), l) <> msg Then GoTo nx
  End If
  FindLabel = i
  Exit Function
nx:
 Next i
 FindLabel = -1 'not found
End Function

Private Function ReadFile(fname As String) As String
 Dim ha As Integer
 On Error GoTo er
 ha = GetAttr(fname) 'err if it doesn't exist
 ha = FreeFile
 Open fname For Binary As #ha
 ReadFile = Space$(LOF(ha))
 Get #ha, , ReadFile
 Close #ha
er:
End Function

Private Sub LoadTranslFile(fname As String, lf As TLangFile)
 Dim e As EErrCode, raw_s As String
 raw_s = ReadFile(fname)
 e = errCannotRead: If Len(raw_s) = 0 Then GoTo er
 ParseTranslation raw_s, TrVersion, lf, False
 lf.FileName = fname

 Exit Sub
er:
 ThrowError e, "LoadTranslation", "Filename: " & fname
End Sub

Private Function ExportTransl(lf As TLangFile) As String
 Dim s As String, i As Long, u As Long
 
 'Headers
 s = lf.Program & vbCrLf & TrVersion & vbCrLf & lf.Language & vbCrLf & lf.Author & vbCrLf

 'Grammatical numbers
 u = UBound(lf.GrammNumbers)
 For i = 0 To u
  With lf.GrammNumbers(i)
   If .def Then
    s = s & "d;"
   Else
    s = s & IIf(.min <> -1, CStr(.min), "") & "," & _
            IIf(.max <> -1, CStr(.max), "") & "," & _
            IIf(.md, CStr(.md), "") & ";"
   End If
  End With
 Next i
 s = Left$(s, Len(s) - 1) & vbCrLf

 'Labels
 For i = 0 To lf.LabelsCt - 1
  s = s & Replace$(lf.Labels(i), vbCrLf, "\" & vbCrLf) & vbCrLf
 Next i
 
 'ret
 ExportTransl = s
End Function

Private Function GetTranslationRelPath(fname As String) As String
 GetTranslationRelPath = Mid$(fname, Len(TargetFolder) + 1)
End Function


'################ Publicly exposed function

Public Sub NewModel(Language As String, GrammNumbers As String, Author As String)
 Dim lf As TLangFile
 lf.Program = App.Title
 lf.Language = Language
 lf.Author = Author
 AnalyzeGN GrammNumbers, lf
 lf.Loaded = True
 p_model = lf 'replace it
End Sub

Public Sub LoadModel(ByVal fname As String, Optional fullpath As Boolean = False)
 If fullpath = False Then fname = TargetFolder & fname
 LoadTranslFile fname, p_model
 p_model.Loaded = True
End Sub

Public Sub LoadTranslation(ByVal fname As String, Optional fullpath As Boolean = False)
 Dim e As EErrCode

 e = errNoLoadedModel
 If p_model.Loaded = False Then GoTo er

 If fullpath = False Then fname = TargetFolder & fname
 LoadTranslFile fname, p_transl
 
 e = errDiffLabelsCount
 If p_model.LabelsCt <> p_transl.LabelsCt Then GoTo er

 p_transl.Loaded = True
 Exit Sub
er:
 ThrowError e, "LoadTranslation", ""
End Sub

Public Function ExportModel() As String
 If p_model.Loaded = False Then ThrowError errNoLoadedModel, "ExportModel"
 ExportModel = ExportTransl(p_model)
End Function

Public Sub UseDefaultTranslation()
 If p_model.Loaded = False Then ThrowError errNoLoadedModel, "UseDefaultTranslation", ""
 p_transl = p_model 'nothing more to do...
End Sub

Public Function Translate(ByVal msg As String, ByVal exact As Boolean, ParamArray var()) As String
'when translation is loaded, compare size!
 Dim v As Variant, i As Long, e As EErrCode

 v = var 'copy the ref (we HAVE to do this to pass it to the Fmt function)

 If Passthrough Then
pt:
  Translate = Fmt(p_model, msg, ThousandsSep, Warn_NotString, DigitsAfterDecimal, NoTrailingZero, v)
  Exit Function
 End If
 
 i = FindLabel(msg, exact, p_model)
 If i = -1 Then
  If AutoInsert Then
   i = p_model.LabelsCt 'Ct-1 is the last one
   ReDim Preserve p_model.Labels(i)
   p_model.LabelsCt = i + 1
   p_model.Labels(i) = msg
   GoTo pt
  Else
   e = errNoTranslation
   GoTo er
  End If 'AutoInsert
 Else
  If AutoInsert Then GoTo pt
 End If 'i = -1

 e = errNoLoadedTranslation
 If p_transl.Loaded = False Then GoTo er

 Translate = Fmt(p_transl, p_transl.Labels(i), ThousandsSep, Warn_NotString, DigitsAfterDecimal, NoTrailingZero, v)
 
 Exit Function
er:
 ThrowError e, "Translate", "Message: """ & msg & """"
End Function

Public Function AskTranslation(parentForm As Form) As String
'Returns a filename or an empty string if nothing was selected.
 Dim fname As String
 TargetFolder = fe.GetCanonicalFolderName(TargetFolder)
 fe.Enumerate TargetFolder, True
 frmTranslations.Show 1, parentForm
 With frmTranslations.lvw
  If .SelectedItem Is Nothing Then GoTo ret
  If .SelectedItem.Selected = False Then GoTo ret
  fname = .SelectedItem.Key
 End With
 AskTranslation = GetTranslationRelPath(fname)
ret:
 Unload frmTranslations
End Function
Private Sub fe_ObjectFound(IsDirectory As Boolean, name As String, sz As Double, ft As String)
 Dim li As ListItem, lf As TLangFile
 
 If IsDirectory Then Exit Sub
 If fe.CheckExtension(name, TargetExtension) = False Then Exit Sub
 'On Error Resume Next
 If ParseTranslation(ReadFile(name), TrVersion, lf, True) = True Then _
  MsgBox Tr.Translate("Warning: an error has occured while parsing ""$""." & vbCrLf & "$", True, name, err.Description), vbExclamation
 On Error GoTo 0
 
 Set li = frmTranslations.lvw.ListItems.Add(, name, lf.Language)
 li.SubItems(1) = lf.Author
End Sub

Public Sub TranslateForm(Frm As Form)
 Dim c As Control, b As Button, ch As ColumnHeader
 For Each c In Frm.Controls
  If (Len(IgnoreTag) = 0) Or (c.Tag <> IgnoreTag) Then
   If (TypeOf c Is Label) Or (TypeOf c Is CommandButton) Or (TypeOf c Is Frame) Or _
      (TypeOf c Is CheckBox) Or (TypeOf c Is OptionButton) Then
    If Len(c.Caption) Then c.Caption = Translate(c.Caption, True)
    If Len(c.ToolTipText) Then c.ToolTipText = Translate(c.ToolTipText, True)
   ElseIf (TypeOf c Is TextBox) Then
    If Len(c.ToolTipText) Then c.ToolTipText = Translate(c.ToolTipText, True)
   ElseIf (TypeOf c Is Menu) Then
    If Len(c.Caption) Then c.Caption = Translate(c.Caption, True)
   ElseIf (TypeOf c Is Toolbar) Then
    For Each b In c.Buttons
     If Len(b.ToolTipText) Then b.ToolTipText = Translate(b.ToolTipText, True)
    Next b
   ElseIf (TypeOf c Is ListView) Then
    For Each ch In c.ColumnHeaders
     If Len(ch.Text) Then ch.Text = Translate(ch.Text, True)
    Next ch
   End If 'controls filter
  End If '(Len(IgnoreTag) = 0) Or (c.Tag <> IgnoreTag)
 Next c
End Sub

Public Property Get CurrentTranslation() As String
 If p_transl.Loaded Then CurrentTranslation = GetTranslationRelPath(p_transl.FileName)
End Property


'################ Constructor

Private Sub Class_Initialize()
 TrVersion = 1
 Warn_NotString = True
 ThousandsSep = True
 DigitsAfterDecimal = -1
 NoTrailingZero = True
 TargetFolder = App.Path & "\lang"
 TargetExtension = "trn"
 IgnoreTag = "dnt" 'do not translate
 Set fe = New clsEnumFiles
End Sub
