Attribute VB_Name = "modHookButton"
Option Explicit
Public Const TME_CANCEL = &H80000000
Public Const TME_HOVER = &H1&
Public Const TME_LEAVE = &H2&
Public Const TME_NONCLIENT = &H10&
Public Const TME_QUERY = &H40000000
Public Const WM_MOUSELEAVE = &H2A3&
Public Type TRACKMOUSEEVENTTYPE
    cbSize As Long
    dwFlags As Long
    hwndTrack As Long
    dwHoverTime As Long
End Type

Public Declare Function TrackMouseEvent Lib "user32" (lpEventTrack As TRACKMOUSEEVENTTYPE) As Long
Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const GWL_WNDPROC = (-4)
Public PrevProc As Long
Public HookedArray As clsArray
Public Sub HookButton(LWnd As Long, ButtonControl As SkinnableButton)
If HookedArray Is Nothing Then Set HookedArray = New clsArray
Dim i As Long
    i = SetWindowLong(LWnd, GWL_WNDPROC, AddressOf WindowProc)
    HookedArray.AddObject LWnd, i, ButtonControl
End Sub
Public Sub UnHookButton(LWnd As Long)
    Dim Prev As Long
    Prev = HookedArray.GetProc(LWnd)
    SetWindowLong LWnd, GWL_WNDPROC, Prev
    HookedArray.RemoveObject LWnd
End Sub
Public Function WindowProc(ByVal hWnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
On Error Resume Next
    Dim Prev As Long
    Prev = HookedArray.GetProc(hWnd)
    If uMsg = WM_MOUSELEAVE Then
        If Prev <> 0 Then
            Dim Obj As SkinnableButton
            Set Obj = HookedArray.GetObj(hWnd)
            Call Obj.HookMsg(uMsg)
        End If
    End If
    WindowProc = CallWindowProc(Prev, hWnd, uMsg, wParam, lParam)
End Function
