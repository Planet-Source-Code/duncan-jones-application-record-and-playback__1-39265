VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cTraceApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'\\ -[cTraceApp]---------------------------------------------------
'\\ Global class to perform application tracing events...
'\\ Requires EventVB_H.dll, from
'\\ http://www.merrioncomputing.com/EventVB/Release_H.htm
'\\ ---------------------------------------------------------------

'\\ Link to the EventVB dll
Private WithEvents vbLink As EventVB.APIFunctions
Attribute vbLink.VB_VarHelpID = -1

'\\ Hooks implemented in that dll
Private WithEvents vbHooks As EventVB.ApiSystemHook
Attribute vbHooks.VB_VarHelpID = -1

'\\ Message window to get callbacks
Private WithEvents vbCallbackWnd As EventVB.ApiWindow
Attribute vbCallbackWnd.VB_VarHelpID = -1

Private sTraceFilename As String
Private nFile As Long

'\\ Different types of tracing....
Private bTraceMouse As Boolean
Private bTraceKeys As Boolean
Private bTraceFocus As Boolean

'\\ Trace modes - playback or record?
Public Enum ApplicationTraceModes
    TM_Record = 1
    TM_Playback = 2
End Enum

Private mMode As ApplicationTraceModes


'\\ API calls to playback events
Private Declare Function SetFocus Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As KeybaordEventFlags, ByVal dwExtraInfo As Long)

Private Declare Function CreateWindowEx Lib "user32" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByVal lpParam As Long) As Long

Public Enum KeybaordEventFlags
    KEYEVENTF_EXTENDEDKEY = &H1
    KEYEVENTF_KEYUP = &H2
End Enum

Public Enum MouseEventFlags
    MOUSEEVENTF_ABSOLUTE = &H8000 '  absolute move - no mickeys required...
    MOUSEEVENTF_LEFTDOWN = &H2 '  left button down
    MOUSEEVENTF_LEFTUP = &H4 '  left button up
    MOUSEEVENTF_MIDDLEDOWN = &H20 '  middle button down
    MOUSEEVENTF_MIDDLEUP = &H40 '  middle button up
    MOUSEEVENTF_MOVE = &H1 '  mouse move
    MOUSEEVENTF_RIGHTDOWN = &H8 '  right button down
    MOUSEEVENTF_RIGHTUP = &H10 '  right button up
End Enum

Private Declare Sub mouse_event Lib "user32" (ByVal dwFlags As MouseEventFlags, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)

Public Enum TraceEventTypes
    TE_MouseButtonDown = &H1
    TE_MouseButtonUp = &H2
    TE_MouseMove = &H3
    TE_MouseWheel = &H4
    TE_KeyDown = &H5
    TE_KeyUp = &H6
    TE_TraceError = &H7
    TE_FocusChange = &H8
End Enum

Private Declare Function GetTickCount Lib "kernel32" () As Long
Private mCreateCount As Long

Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function GetFocus Lib "user32" () As Long
Private Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long

Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32" () As Long

Private Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SetCursorPos Lib "user32" (ByVal x As Long, ByVal y As Long) As Long

Private Declare Function LPtoDP Lib "gdi32" (ByVal hdc As Long, lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Sub DecodeFocusChange(ByVal EventLine As String, hwndTarget As Long)

Dim sTarget As String
If InStr(EventLine, "Target=") Then
    sTarget = Mid$(EventLine, InStr(EventLine, "Target=") + Len("Target="))
    If InStr(sTarget, ",") Then
        sTarget = Left$(sTarget, InStr(sTarget, ",") - 1)
    End If
    If InStr(sTarget, ":") Then
        sTarget = Left$(sTarget, InStr(sTarget, ":") - 1)
    End If
    hwndTarget = MCL_TracingUtilities.ControlNameToWindowhandle(sTarget)
End If

End Sub

Private Sub DecodeKeyboardEvent(ByVal EventLine As String, bVk As Byte, bScan As Byte, dwFlags As Long, dwExtraInfo As Long, hwndTarget As Long, msgTime As Long)

'\\ EventLine - 5:Key Down:Key=160,Alt=False,Extended=False,Target=Form1.Text1:Time=3836
If Left$(EventLine, 1) = TE_KeyUp Then
    dwFlags = KEYEVENTF_KEYUP
End If

If InStr(EventLine, "Key=") Then
    bVk = Val(Mid$(EventLine, InStr(EventLine, "Key=") + Len("Key=")))
End If

If InStr(EventLine, "Time=") Then
    msgTime = Val(Mid$(EventLine, InStr(EventLine, "Time=") + Len("Time=")))
End If

'\\ Get the target for the mouse event - if any exists
Dim sTarget As String
If InStr(EventLine, "Target=") Then
    sTarget = Mid$(EventLine, InStr(EventLine, "Target=") + Len("Target="))
    If InStr(sTarget, ",") Then
        sTarget = Left$(sTarget, InStr(sTarget, ",") - 1)
    End If
    If InStr(sTarget, ":") Then
        sTarget = Left$(sTarget, InStr(sTarget, ":") - 1)
    End If
    hwndTarget = MCL_TracingUtilities.ControlNameToWindowhandle(sTarget)
End If

End Sub

Private Sub DecodeMouseEvent(ByVal EventLine As String, dFlags As MouseEventFlags, dx As Long, dy As Long, cButtons As Long, ExtraInfo As Long, msgTime As Long, hwndTarget As Long)

'\\ EventLine - 2:Mouse Button Up:Button=1,x=240,y=229,Target=Form1.Text1:Time=2964
Dim lButton As Long
If InStr(EventLine, "Button=") Then
    lButton = Val(Mid$(EventLine, InStr(EventLine, "Button=") + Len("Button=")))
End If


Select Case Val(Left$(EventLine, 1))
Case TE_MouseButtonDown
    If lButton = vbLeftButton Then
        dFlags = MOUSEEVENTF_LEFTDOWN
    ElseIf lButton = vbMiddleButton Then
        dFlags = MOUSEEVENTF_MIDDLEDOWN
    Else
        dFlags = MOUSEEVENTF_RIGHTDOWN
    End If
Case TE_MouseButtonUp
    If lButton = vbLeftButton Then
        dFlags = MOUSEEVENTF_LEFTUP
    ElseIf lButton = vbMiddleButton Then
        dFlags = MOUSEEVENTF_MIDDLEUP
    Else
        dFlags = MOUSEEVENTF_RIGHTUP
    End If
Case TE_MouseMove
    dFlags = MOUSEEVENTF_MOVE
End Select

If InStr(EventLine, "x=") Then
    dx = Val(Mid$(EventLine, InStr(EventLine, "x=") + Len("x=")))
End If

If InStr(EventLine, "y=") Then
    dy = Val(Mid$(EventLine, InStr(EventLine, "y=") + Len("y=")))
End If

If InStr(EventLine, "Time=") Then
    msgTime = Val(Mid$(EventLine, InStr(EventLine, "Time=") + Len("Time=")))
End If

'\\ Get the target for the mouse event - if any exists
Dim sTarget As String
If InStr(EventLine, "Target=") Then
    sTarget = Mid$(EventLine, InStr(EventLine, "Target=") + Len("Target="))
    If InStr(sTarget, ",") Then
        sTarget = Left$(sTarget, InStr(sTarget, ",") - 1)
    End If
    If InStr(sTarget, ":") Then
        sTarget = Left$(sTarget, InStr(sTarget, ":") - 1)
    End If
    hwndTarget = MCL_TracingUtilities.ControlNameToWindowhandle(sTarget)
End If

'\\ All mouse events are absolute
dFlags = MOUSEEVENTF_ABSOLUTE Or dFlags

'\\ and must be offset by the position of the target window
Dim wndOffset As New ApiWindow
wndOffset.hwnd = hwndTarget
dx = dx + wndOffset.RECT.Left
dy = dy + wndOffset.RECT.Top

End Sub

Public Property Get Mode() As ApplicationTraceModes

    Mode = mMode

End Property

Public Property Let Mode(ByVal newval As ApplicationTraceModes)

If mMode <> newval Then
    mMode = newval
End If

End Property

Private Function OffsetLocation(ByVal Location As APIPoint, ByVal hwnd As ApiWindow)

Location.x = Location.x - hwnd.RECT.Left
Location.y = Location.y - hwnd.RECT.Top

End Function

'\\ --[Playback]-------------------------------------------------------------
'\\ Reads events from the trace log and plays them back to the application
'\\ -------------------------------------------------------------------------
Private Sub PlayBack()

Dim sEventLine As String

Dim hwndTarget As Long
'\\ Parameters for a keyboard messge
Dim bVk As Byte, bScan As Byte, dwFlags As Long, dwExtraInfo As Long
'\\ parameters for a mouse message
Dim dFlags As MouseEventFlags, dx As Long, dy As Long, cButtons As Long, ExtraInfo As Long
Dim lMsgTime As Long, msgType As TraceEventTypes

Static hwndKeyTarget As Long

If Not EOF(nFile) Then
    '\\ Read next event
    Line Input #nFile, sEventLine
    msgType = Left$(sEventLine, 1)
    Select Case msgType
    Case TE_TraceError
        Debug.Print "An error occured recording the trace information"
        Stop
    Case TE_MouseButtonDown, TE_MouseButtonUp
        '\\ A mouse event occured - decode it
        Call DecodeMouseEvent(sEventLine, dFlags, dx, dy, cButtons, ExtraInfo, lMsgTime, hwndTarget)
        '\\ wait until the app is ready for input...
        While (GetTickCount - mCreateCount) < lMsgTime
            DoEvents
        Wend
        '\\ Send the event to the app
        mouse_event dFlags, dx, dy, cButtons, ExtraInfo
        
    Case TE_MouseMove
        '\\ A mouse move event occured - decode it
        Call DecodeMouseEvent(sEventLine, dFlags, dx, dy, cButtons, ExtraInfo, lMsgTime, hwndTarget)
        '\\ Put the cursor at the new position
        SetCursorPos dx, dy
        
    Case TE_KeyDown, TE_KeyUp
        '\\ A keyboard evenyt occured - decode it
        Call DecodeKeyboardEvent(sEventLine, bVk, bScan, dwFlags, dwExtraInfo, hwndTarget, lMsgTime)
        '\\ wait until the app is ready for input...
        While (GetTickCount - mCreateCount) < lMsgTime
            DoEvents
        Wend
        If hwndTarget <> hwndKeyTarget Then
            hwndKeyTarget = hwndTarget
            Call SetFocus(hwndTarget)
        End If
        '\\ Send the event to the app
        keybd_event bVk, bScan, dwFlags, dwExtraInfo

    Case TE_FocusChange
        Call DecodeFocusChange(sEventLine, hwndTarget)
        If hwndTarget <> 0 Then
            Call SetFocus(hwndTarget)
        End If
        
    End Select

    '\\ Allow this event to be processed
    DoEvents
    
    '\\ Ask for the next record
    '\\ Note that we post the message rather than using Sendmessage so that other events
    '\\ which occur in our program may be processed
    PostMessage vbCallbackWnd.hwnd, MCL_TracingUtilities.NextTraceRecord, 0, 0
Else
    '\\ End of playback reached...
    Close #nFile
End If

End Sub

Private Sub SaveEvent(ByVal EventType As TraceEventTypes, ByVal EventData As String)

Static mPrevTime As Long
Dim lCurrTime

lCurrTime = GetTickCount - mCreateCount
If mPrevTime <> lCurrTime Then
    mPrevTime = lCurrTime
    Select Case EventType
    Case TE_MouseButtonDown
        Print #nFile, EventType & ":Mouse button down:" & EventData & ":Time=" & lCurrTime
    Case TE_MouseButtonUp
        Print #nFile, EventType & ":Mouse Button Up:" & EventData & ":Time=" & lCurrTime
    Case TE_MouseMove
        Print #nFile, EventType & ":Mouse Move:" & EventData & ":Time=" & lCurrTime
    Case TE_MouseWheel
        Print #nFile, EventType & ":Mouse Wheel:" & EventData & ":Time=" & lCurrTime
    Case TE_KeyDown
        Print #nFile, EventType & ":Key Down:" & EventData & ":Time=" & lCurrTime
    Case TE_KeyUp
        Print #nFile, EventType & ":Key Up:" & EventData & ":Time=" & lCurrTime
    Case TE_TraceError
        Print #nFile, EventType & ":**ERROR**:" & EventData
    Case TE_FocusChange
        Print #nFile, EventType & ":Focus Changed:" & EventData & ":Time=" & lCurrTime
    End Select
End If

End Sub


Public Property Let TraceFocusChanges(ByVal newval As Long)

If newval <> bTraceFocus Then
    If Not newval Then
        '\\ If mouse hook is running, disable it
        If vbHooks.HookIdByType(WH_CBT) <> 0 Then
            vbHooks.StopHook WH_CBT
        End If
    Else
        '\\ If the mouse hook is not running, start it
        If vbHooks.HookIdByType(WH_CBT) = 0 And Mode = TM_Record Then
            vbHooks.StartHook WH_CBT, HOOK_LOCAL_PROCESS
        End If
    End If
    bTraceKeys = newval
End If

End Property

Public Property Let TraceKeyboardEvents(ByVal newval As Boolean)

If newval <> bTraceKeys Then
    If Not newval Then
        '\\ If mouse hook is running, disable it
        If vbHooks.HookIdByType(WH_KEYBOARD_LL) <> 0 Then
            vbHooks.StopHook WH_KEYBOARD_LL
        End If
    Else
        '\\ If the mouse hook is not running, start it
        If vbHooks.HookIdByType(WH_KEYBOARD_LL) = 0 And Mode = TM_Record Then
            vbHooks.StartHook WH_KEYBOARD_LL, HOOK_GLOBAL
        End If
    End If
    bTraceKeys = newval
End If

End Property

Public Property Let TraceMouseEvents(ByVal newval As Boolean)

If newval <> bTraceMouse Then
    If Not newval Then
        '\\ If mouse hook is running, disable it
        If vbHooks.HookIdByType(WH_MOUSE) <> 0 Then
            vbHooks.StopHook WH_MOUSE
        End If
    Else
        '\\ If the mouse hook is not running, start it
        If vbHooks.HookIdByType(WH_MOUSE) = 0 And Mode = TM_Record Then
            vbHooks.StartHook WH_MOUSE, HOOK_LOCAL_PROCESS
        End If
    End If
    bTraceMouse = newval
End If

End Property


Private Sub Class_Initialize()

'\\ Time when ticker was created....
mCreateCount = GetTickCount

Set vbLink = New EventVB.APIFunctions
Set vbHooks = vbLink.System.Hooks

    Set vbCallbackWnd = New ApiWindow
    vbCallbackWnd.hwnd = CreateWindowEx(0, "STATIC", "MCL Trace Client 1.1", 0, 0, 0, 0, 0, 0, 0, 0, 0)
    
    '\\ Subclass the callback window top get messages
    vbLink.SubclassedWindows.Add vbCallbackWnd


End Sub

Public Property Let Filename(ByVal newval As String)

If newval <> sTraceFilename Then
    '\\ If changeing, close the old file
    If nFile <> 0 Then
        Close #nFile
        nFile = 0
    End If
    If newval <> "" Then
        sTraceFilename = newval
        nFile = FreeFile
        If Mode = TM_Record Then
            Open sTraceFilename For Output As nFile
        Else
            Open sTraceFilename For Input As nFile
            Call PlayBack
        End If
    End If
End If

End Property

Private Sub Class_Terminate()

'\\ If a file is open, close it...
If nFile <> 0 Then
    Close #nFile
End If
'\\ Close and destroy the callback window
If Not vbCallbackWnd Is Nothing Then
    vbCallbackWnd.SendMessage WM_CLOSE, 0, 0
    vbLink.SubclassedWindows.Remove vbCallbackWnd
    Set vbCallbackWnd = Nothing
End If
'\\ Unlink any running hooks
vbHooks.StopHook WH_KEYBOARD_LL
vbHooks.StopHook WH_MOUSE
vbHooks.StopHook WH_CBT

Set vbHooks = Nothing
'\\ Unlink the EventVB.dll
Set vbLink = Nothing

End Sub

Private Sub vbCallbackWnd_WindowMessageFired(ByVal msg As EventVB.WindowMessages, ByVal wParam As Long, ByVal lParam As Long, Cancel As Boolean, ProcRet As Long)

If msg = MCL_TracingUtilities.NextTraceRecord Then
    '\\ Get the next record and action it
    Call PlayBack
End If

End Sub


Private Sub vbHooks_BeforeFocusChanged(ByVal GainFocus As EventVB.ApiWindow, ByVal LoseFocus As EventVB.ApiWindow, Cancel As Boolean)

SaveEvent TE_FocusChange, "From=" & WindowhandleToControlName(LoseFocus) & ",To=" & WindowhandleToControlName(GainFocus)

End Sub

Private Sub vbHooks_KeyDown(ByVal VKey As Long, ByVal scanCode As Long, ByVal ExtendedKey As Boolean, ByVal AltDown As Boolean, ByVal Injected As Boolean, Cancel As Boolean)

'\\ We are only interested in keyboard events for this application
If GetFocus <> 0 Then
    SaveEvent TE_KeyDown, "Key=" & VKey & ",Alt=" & AltDown & ",Extended=" & ExtendedKey & ",Target=" & WindowhandleToControlName(GetFocus)
End If

End Sub

Private Sub vbHooks_KeyUp(ByVal VKey As Long, ByVal scanCode As Long, ByVal ExtendedKey As Boolean, ByVal AltDown As Boolean, ByVal Injected As Boolean, Cancel As Boolean)

'\\ We are only interested in keyboard events for this application
If GetFocus <> 0 Then
    SaveEvent TE_KeyUp, "Key=" & VKey & ",Alt=" & AltDown & ",Extended=" & ExtendedKey & ",Target=" & WindowhandleToControlName(GetFocus)
End If

End Sub


Private Sub vbHooks_MouseButtonDown(ByVal Button As Integer, ByVal Location As EventVB.APIPoint, ByVal TargetWindow As EventVB.ApiWindow, Cancel As Boolean)

'\\ Mouse events a recorded relative to the screen - but this may change between
'\\ recorder and playback so it is better to offset them to relative to the window dc
Call OffsetLocation(Location, TargetWindow)
SaveEvent TE_MouseButtonDown, "Button=" & Button & ",x=" & Location.x & ",y=" & Location.y & ",Target=" & WindowhandleToControlName(TargetWindow.hwnd)

End Sub


Private Sub vbHooks_MouseButtonUp(ByVal Button As Integer, ByVal Location As EventVB.APIPoint, ByVal TargetWindow As EventVB.ApiWindow, Cancel As Boolean)

'\\ Mouse events a recorded relative to the screen - but this may change between
'\\ recorder and playback so it is better to offset them to relative to the window dc
Call OffsetLocation(Location, TargetWindow)
SaveEvent TE_MouseButtonUp, "Button=" & Button & ",x=" & Location.x & ",y=" & Location.y & ",Target=" & WindowhandleToControlName(TargetWindow.hwnd)

End Sub


Private Sub vbHooks_MouseMove(ByVal Location As EventVB.APIPoint, ByVal TargetWindow As EventVB.ApiWindow, Cancel As Boolean)

'\\ For a journal hook, mouse move is always sent to hwnd=0,
'\\ which is not useful to us, so we need to look at the
'\\ window underneath the point

Dim hwndTarget As New ApiWindow
hwndTarget.hwnd = WindowFromPoint(Location.x, Location.y)

If hwndTarget.Thread.ThreadID = vbLink.CurrentThreadId Then
    Call OffsetLocation(Location, hwndTarget)
    SaveEvent TE_MouseMove, "x=" & Location.x & ",y=" & Location.y & ",Target=" & WindowhandleToControlName(hwndTarget.hwnd)
End If

End Sub


Private Sub vbLink_ApiError(ByVal Number As Long, ByVal Source As String, ByVal Description As String)

If Mode = TM_Record Then
    SaveEvent TE_TraceError, Description & " in " & Source
Else
    Debug.Print "Error : " & Description & " in " & Source
    Stop
End If

End Sub

