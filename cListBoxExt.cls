VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cListBoxExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const MODULE_NAME As String = "cListBoxExt"

'=========================================================================
' Events
'=========================================================================

Event ItemCheck(Item As Integer)

'=========================================================================
' API
'=========================================================================

'--- windows messages
Private Const WM_DESTROY                    As Long = &H2
Private Const WM_PAINT                      As Long = &HF
Private Const WM_DRAWITEM                   As Long = &H2B
Private Const WM_CTLCOLORLISTBOX            As Long = &H134
Private Const WM_LBUTTONDOWN                As Long = &H201
Private Const WM_LBUTTONDBLCLK              As Long = &H203
Private Const WM_PRINTCLIENT                As Long = &H318
'--- for GetDeviceCaps
Private Const LOGPIXELSY                    As Long = 90 ' Logical pixels/inch in Y
'--- for CreateFont
Private Const FW_NORMAL                     As Long = 400
'--- for ExtTextOut
Private Const ETO_OPAQUE                    As Long = 2
Private Const ETO_CLIPPED                   As Long = 4

Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function CreateFont Lib "gdi32" Alias "CreateFontA" (ByVal nHeight As Long, ByVal nWidth As Long, ByVal nEscapement As Long, ByVal nOrientation As Long, ByVal fnWeigh As Long, ByVal fdwItalic As Long, ByVal fdwUnderline As Long, ByVal fdwStrikeOut As Long, ByVal fdwCharSet As Long, ByVal fdwOutputPrecision As Long, ByVal fdwClipPrecision As Long, ByVal fdwQuality As Long, ByVal fdwPitchAndFamily As Long, ByVal lpszFace As String) As Long
Private Declare Function ExtTextOut Lib "gdi32" Alias "ExtTextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal wOptions As Long, lpRect As RECT, ByVal lpString As String, ByVal nCount As Long, ByVal lpDx As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetBkColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long

Private Type RECT
    Left            As Long
    Top             As Long
    Right           As Long
    Bottom          As Long
End Type

Private Type DRAWITEMSTRUCT
    CtlType         As Long
    CtlID           As Long
    itemID          As Long
    itemAction      As Long
    itemState       As Long
    hwndItem        As Long
    hdc             As Long
    rcItem          As RECT
    itemData        As Long
End Type

'=========================================================================
' Constants and member variables
'=========================================================================

Private Const PREFIX_GREY       As String = "#"
Private Const PREFIX_CANGREY    As String = "C"

Private WithEvents m_oCtl   As ListBox
Attribute m_oCtl.VB_VarHelpID = -1
Private m_hWnd              As Long
Private m_uSubclass         As SubClassData
Private m_uSubclassParent   As SubClassData
Private m_cGrey             As Collection
Private m_bInSet            As Boolean
Private m_bGreyedNotChecked As Boolean
Private m_lLastItem         As Long

'=========================================================================
' Error handling
'=========================================================================

Private Sub PrintError(sFunc As String)
    Debug.Print MODULE_NAME & "." & sFunc & ": " & Err.Description
End Sub

'=========================================================================
' Properties
'=========================================================================

Property Get Selected(ByVal Index As Integer) As CheckBoxConstants
    If pvGrey(PREFIX_GREY, Index) Then
        Selected = vbGrayed
    Else
        Selected = -m_oCtl.Selected(Index)
    End If
End Property

Property Let Selected(ByVal Index As Integer, ByVal eValue As CheckBoxConstants)
    m_bInSet = True
    m_oCtl.Selected(Index) = (eValue <> vbUnchecked)
    pvGrey(PREFIX_GREY, Index) = (eValue = vbGrayed)
    CanGrey(Index) = CanGrey(Index) Or (eValue = vbGrayed)
    InvalidateRect m_hWnd, 0, 0
    m_bInSet = False
End Property

Property Get CanGrey(ByVal Index As Integer) As Boolean
    CanGrey = pvGrey(PREFIX_CANGREY, Index)
End Property

Property Let CanGrey(ByVal Index As Integer, ByVal bValue As Boolean)
    pvGrey(PREFIX_CANGREY, Index) = bValue
End Property

Property Get GreyedNotChecked() As Boolean
    GreyedNotChecked = m_bGreyedNotChecked
End Property

Property Let GreyedNotChecked(ByVal bValue As Boolean)
    m_bGreyedNotChecked = bValue
    InvalidateRect m_hWnd, 0, 0
End Property

'= private ===============================================================

Property Get pvGrey(sPrefix As String, ByVal Index As Integer) As Boolean
    On Error Resume Next
    If Not m_cGrey.Item(sPrefix & Index) Then
    Else
        pvGrey = True
    End If
    On Error GoTo 0
End Property

Property Let pvGrey(sPrefix As String, ByVal Index As Integer, ByVal bValue As Boolean)
    On Error Resume Next
    If bValue Then
        m_cGrey.Add True, sPrefix & Index
    Else
        m_cGrey.Remove sPrefix & Index
    End If
    On Error GoTo 0
End Property

'=========================================================================
' Methods
'=========================================================================

Public Function Init(oCtl As Object, Optional bGreyedNotChecked As Boolean)
    '--- init member vars
    Set m_oCtl = oCtl
    m_bGreyedNotChecked = bGreyedNotChecked
    Set m_cGrey = New Collection
    m_lLastItem = -1
    '--- subclass control
    m_hWnd = m_oCtl.hwnd
    SubClass m_uSubclass, m_hWnd, ObjPtr(Me), AddressOf RedirectListBoxWndProc
End Function

Public Sub Terminate()
    '--- unsubclass and cleanup
    UnSubClass m_uSubclass, m_hWnd
    Set m_oCtl = Nothing
End Sub

'= friend ================================================================

Friend Function frWndProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Const FUNC_NAME     As String = "frWndProc"
    Dim hWndParent      As Long
    
    On Error GoTo EH
    Select Case wMsg
    Case WM_PAINT, WM_PRINTCLIENT '--- WM_PRINTCLIENT is in case MS fixes the XP/w2k bug
        hWndParent = GetParent(m_hWnd)
        If hWndParent <> 0 Then
            SubClass m_uSubclassParent, hWndParent, ObjPtr(Me), AddressOf RedirectListBoxParentWndProc
        End If
        frWndProc = CallWindowProc(m_uSubclass.WndProcNext, hwnd, wMsg, wParam, lParam)
        If hWndParent <> 0 Then
            UnSubClass m_uSubclassParent, hWndParent
        End If
    Case WM_LBUTTONDOWN, WM_LBUTTONDBLCLK
        '--- set horizontal position to 0
        lParam = (lParam And &HFFFF0000)
        '--- note: dbl-click is forwarded as single-click too
        frWndProc = CallWindowProc(m_uSubclass.WndProcNext, hwnd, WM_LBUTTONDOWN, wParam, lParam)
    Case WM_DESTROY
        For m_lLastItem = m_oCtl.ListCount - 1 To 0 Step -1
            If m_oCtl.Selected(m_lLastItem) Then
                Exit For
            End If
        Next
        hWndParent = GetParent(m_hWnd)
        UnSubClass m_uSubclass, m_hWnd
        m_hWnd = 0
        SubClass m_uSubclassParent, hWndParent, ObjPtr(Me), AddressOf RedirectListBoxParentWndProc
    Case Else
        frWndProc = CallWindowProc(m_uSubclass.WndProcNext, hwnd, wMsg, wParam, lParam)
    End Select
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

Friend Function frParentWndProc(ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    Const FUNC_NAME     As String = "frParentWndProc"
    Const RECT_SIZE     As Long = 9
    Const LEFT_OFFESET  As Long = 3
    Const FONT_SIZE     As Long = 13
    Const FONT_NAME     As String = "Marlett"
    Dim dis             As DRAWITEMSTRUCT
    Dim rc              As RECT
    Dim hPrevFont       As Long
    Dim clrPrevBack     As Long
    Dim clrPrevFore     As Long
    Dim sText           As String
    
    On Error GoTo EH
    Select Case wMsg
    Case WM_DRAWITEM
        '--- dereference struct
        CopyMemory dis, ByVal lParam, Len(dis)
        '--- call original wndproc
        frParentWndProc = CallWindowProc(m_uSubclassParent.WndProcNext, hwnd, wMsg, wParam, lParam)
        '--- draw greyed checkmark
        If pvGrey(PREFIX_GREY, dis.itemID) Then
            rc = dis.rcItem
            With rc
                .Left = .Left + LEFT_OFFESET
                .Top = .Top + ((.Bottom - .Top) - RECT_SIZE) \ 2
                .Bottom = .Top + RECT_SIZE
                .Right = .Left + RECT_SIZE
            End With
            hPrevFont = SelectObject(dis.hdc, CreateFont(FONT_SIZE, 0, 0, 0, FW_NORMAL, 0, 0, 0, 1, 0, 0, 0, 0, FONT_NAME))
            clrPrevBack = SetBkColor(dis.hdc, GetSysColor(vbButtonFace And &HFF))
            clrPrevFore = SetTextColor(dis.hdc, GetSysColor(IIf(m_oCtl.Enabled, vbWindowText, vbButtonShadow) And &HFF))
            If Not m_bGreyedNotChecked Then
                '--- this shows as checkmark when drawn using Marlett font
                sText = "b"
            End If
            ExtTextOut dis.hdc, rc.Left - 2, rc.Top - 2, ETO_OPAQUE Or ETO_CLIPPED, rc, sText, Len(sText), 0
            Call DeleteObject(SelectObject(dis.hdc, hPrevFont))
            Call SetBkColor(dis.hdc, clrPrevBack)
            Call SetTextColor(dis.hdc, clrPrevFore)
        End If
    Case WM_CTLCOLORLISTBOX
        If m_hWnd = 0 Then
            m_hWnd = m_oCtl.hwnd
            SubClass m_uSubclass, m_hWnd, ObjPtr(Me), AddressOf RedirectListBoxWndProc
            UnSubClass m_uSubclassParent, hwnd
        End If
    Case Else
        frParentWndProc = CallWindowProc(m_uSubclassParent.WndProcNext, hwnd, wMsg, wParam, lParam)
    End Select
    Exit Function
EH:
    PrintError FUNC_NAME
    Resume Next
End Function

'=========================================================================
' Control events
'=========================================================================

Private Sub m_oCtl_ItemCheck(Item As Integer)
    Const FUNC_NAME     As String = "m_oCtl_ItemCheck"
    
    On Error GoTo EH
    If m_lLastItem > 0 Then
        If Item = m_lLastItem Then
            m_lLastItem = -1
        End If
        Exit Sub
    End If
    If Not m_bInSet Then
        If Selected(Item) = vbUnchecked And CanGrey(Item) Then
            m_oCtl.Selected(Item) = True
            pvGrey(PREFIX_GREY, Item) = True
        Else
            pvGrey(PREFIX_GREY, Item) = False
        End If
        RaiseEvent ItemCheck(Item)
    End If
    Exit Sub
EH:
    PrintError FUNC_NAME
    Resume Next
End Sub
