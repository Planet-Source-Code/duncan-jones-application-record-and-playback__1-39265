Attribute VB_Name = "MCL_TracingUtilities"
Option Explicit

Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, lpParam As Any) As Long

Private lOldWndproc As Long

Public Function ControlNameToWindowhandle(ByVal sControl As String) As Long

'\\ scontrol is in the format form.control[(index)]
Dim FormName As String
Dim ControlName As String
Dim IndexPart As String

FormName = Left$(sControl, InStr(sControl, ".") - 1)
If InStr(sControl, "(") Then
    ControlName = Mid$(sControl, InStr(sControl, ".") + 1)
    IndexPart = Mid$(ControlName, InStr(ControlName, "("))
    ControlName = Left$(ControlName, Len(ControlName) - Len(IndexPart))
    IndexPart = Mid$(IndexPart, 2)
    IndexPart = Left$(IndexPart, Len(IndexPart) - 1)
Else
    ControlName = Mid$(sControl, InStr(sControl, ".") + 1)
End If

Dim fThis As Form

For Each fThis In Forms
    If fThis.Name = FormName Then
        If ControlName <> "" Then
            If IndexPart = "" Then
                ControlNameToWindowhandle = fThis.Controls(ControlName).hwnd
            Else
                ControlNameToWindowhandle = fThis.Controls(ControlName).Item(IndexPart).hwnd
            End If
        Else
            ControlNameToWindowhandle = fThis.hwnd
        End If
        Exit For
    End If
Next fThis

End Function

Public Function NextTraceRecord() As Long

Static lNextRecord As Long

If lNextRecord = 0 Then
    lNextRecord = RegisterWindowMessage("MCL Trace:Next")
End If

NextTraceRecord = lNextRecord

End Function

Private Function ParentForm(ByVal hwnd As Long) As Form
   Dim i%, pWindow&, fWindow&, fThis As Form
    
   pWindow = GetParent(hwnd)
   While pWindow
      fWindow = pWindow
      pWindow = GetParent(pWindow)
   Wend
       
   For Each fThis In Forms
      If (fThis.hwnd = fWindow) Or (fThis.hwnd = hwnd) Then
            Set ParentForm = fThis
            Exit For
      End If
   Next fThis
   
  
End Function




Public Function WindowhandleToControlName(ByVal hwnd As Long) As String

Dim fThis As Form
Dim ctThis As Control

Dim FormName As String, ControlName As String

Set fThis = ParentForm(hwnd)

If Not fThis Is Nothing Then
    If fThis.hwnd = hwnd Then
        FormName = fThis.Name
    Else
        For Each ctThis In fThis.Controls
            If Not (TypeOf ctThis Is Label) Or (TypeOf ctThis Is Shape) Or (TypeOf ctThis Is Image) Then
              If ctThis.hwnd = hwnd Then
                FormName = fThis.Name
                If TypeName(fThis.Controls(ctThis.Name)) = "Object" Then
                    ControlName = ctThis.Name & "(" & ctThis.Index & ")"
                Else
                   ControlName = ctThis.Name
                End If
              End If
            End If
        Next ctThis
    End If
End If

WindowhandleToControlName = FormName & "." & ControlName

End Function

Public Function SaveTraceSetting(ByVal bTrace As Boolean, ByVal Filename As String, ByVal Keyboard As Boolean, ByVal Mouse As Boolean, ByVal Focus As Boolean)

Call SaveSetting(App.EXEName, "Tracing", "Trace", IIf(bTrace, "1", "0"))
Call SaveSetting(App.EXEName, "Tracing", "Filename", Filename)
Call SaveSetting(App.EXEName, "Tracing", "Mouse", IIf(Mouse, "1", "0"))
Call SaveSetting(App.EXEName, "Tracing", "Keyboard", IIf(Keyboard, "1", "0"))
Call SaveSetting(App.EXEName, "Tracing", "Focus", IIf(Focus, "1", "0"))

End Function
