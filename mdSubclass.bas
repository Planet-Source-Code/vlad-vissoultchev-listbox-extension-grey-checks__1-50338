Attribute VB_Name = "mdSubclass"
'==============================================================================
' From Matt Curland's Advanced Visual Basic (www.powervb.com)
'==============================================================================
Option Explicit

'==============================================================================
' API
'==============================================================================

'--- for Get/SetWindowLong
Private Const GWL_WNDPROC               As Long = -4

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Private Declare Function SetTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
'Private Declare Function KillTimer Lib "user32" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
'Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long

Private Type ThunkBytes
    Thunk(5)                As Long
End Type

Public Type PushParamThunk
    pfn                     As Long
    Code                    As ThunkBytes
End Type

Public Type SubClassData
    WndProcNext             As Long
    WndProcThunkThis        As PushParamThunk
    #If DEBUGWINDOWPROC Then
        dbg_Hook            As WindowProcHook
    #End If
End Type

Public Sub InitPushParamThunk(Thunk As PushParamThunk, ByVal ParamValue As Long, ByVal pfnDest As Long)
'push [esp]
'mov eax, 16h // Dummy value for parameter value
'mov [esp + 4], eax
'nop // Adjustment so the next long is nicely aligned
'nop
'nop
'mov eax, 1234h // Dummy value for function
'jmp eax
'nop
'nop
    
    With Thunk.Code
        .Thunk(0) = &HB82434FF
        .Thunk(1) = ParamValue
        .Thunk(2) = &H4244489
        .Thunk(3) = &HB8909090
        .Thunk(4) = pfnDest
        .Thunk(5) = &H9090E0FF
    End With
    Thunk.pfn = VarPtr(Thunk.Code)
End Sub

Public Sub SubClass(Data As SubClassData, ByVal hWnd As Long, ByVal ThisPtr As Long, ByVal pfnRedirect As Long)
    With Data
        If .WndProcNext Then
            SetWindowLong hWnd, GWL_WNDPROC, .WndProcNext
            .WndProcNext = 0
        End If
        InitPushParamThunk .WndProcThunkThis, ThisPtr, pfnRedirect
#If DEBUGWINDOWPROC Then
        On Error Resume Next
        Set .dbg_Hook = Nothing
        Set .dbg_Hook = CreateWindowProcHook
        If Err Then
            On Error GoTo 0
            Exit Sub
        End If
        On Error GoTo 0
        With .dbg_Hook
            .SetMainProc Data.WndProcThunkThis.pfn
            Data.WndProcNext = SetWindowLong(hWnd, GWL_WNDPROC, .ProcAddress)
            .SetDebugProc Data.WndProcNext
        End With
#Else
        .WndProcNext = SetWindowLong(hWnd, GWL_WNDPROC, .WndProcThunkThis.pfn)
#End If
    End With
End Sub

Public Sub UnSubClass(Data As SubClassData, ByVal hWnd As Long)
    With Data
        If .WndProcNext Then
            SetWindowLong hWnd, GWL_WNDPROC, .WndProcNext
            .WndProcNext = 0
        End If
#If DEBUGWINDOWPROC Then
        Set .dbg_Hook = Nothing
#End If
    End With
End Sub

'==============================================================================
' Sample redirectors
'==============================================================================

'Public Function RedirectControlWndProc( _
'            ByVal This As MyControl, _
'            ByVal hWnd As Long, _
'            ByVal uMsg As Long, _
'            ByVal wParam As Long, _
'            ByVal lParam As Long) As Long
'    Select Case uMsg
'    Case WM_CANCELMODE
'        This.frCancelMode
'    End Select
'    ControlWndProc = CallWindowProc(This.frGetWndProcNext, hWnd, uMsg, wParam, ByVal lParam)
'End Function

