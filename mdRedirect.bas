Attribute VB_Name = "mdRedirect"
Option Explicit

Public Function RedirectListBoxWndProc( _
            ByVal This As cListBoxExt, _
            ByVal hwnd As Long, _
            ByVal uMsg As Long, _
            ByVal wParam As Long, _
            ByVal lParam As Long) As Long
    RedirectListBoxWndProc = This.frWndProc(hwnd, uMsg, wParam, lParam)
End Function

Public Function RedirectListBoxParentWndProc( _
            ByVal This As cListBoxExt, _
            ByVal hwnd As Long, _
            ByVal uMsg As Long, _
            ByVal wParam As Long, _
            ByVal lParam As Long) As Long
    RedirectListBoxParentWndProc = This.frParentWndProc(hwnd, uMsg, wParam, lParam)
End Function

