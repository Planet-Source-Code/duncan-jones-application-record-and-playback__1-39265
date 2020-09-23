Attribute VB_Name = "TraceTestMain"
Option Explicit

Dim oTracing As cTraceApp

Public Sub main()

Dim bTrace As Boolean

'\\ To debug the tracing element, uncomment the following line
Call SaveTraceSetting(True, "MCLTrace.txt", True, True, True)

bTrace = GetSetting(App.EXEName, "Tracing", "Trace", "0")
'\\ If tracing is needed, create a new cTraceApp object
If bTrace Then
    Set oTracing = New cTraceApp
    With oTracing
        '.Mode = TM_Record '\\ Set this to TM_Playback in the debug environment
                          '\\ to playback the file
        .Mode = TM_Playback
        .Filename = GetSetting(App.EXEName, "Tracing", "Filename", "MCLTrace.txt")
        .TraceKeyboardEvents = GetSetting(App.EXEName, "Tracing", "Keyboard", "0")
        .TraceMouseEvents = GetSetting(App.EXEName, "Tracing", "Mouse", "0")
        .TraceFocusChanges = GetSetting(App.EXEName, "Tracing", "Focus", "0")
    End With
End If

Form1.Show vbModal

'\\ Stop tracing once you are done....
If bTrace Then
    Set oTracing = Nothing
End If

End Sub




