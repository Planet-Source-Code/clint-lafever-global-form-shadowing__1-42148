Attribute VB_Name = "basCOMMON"
Option Explicit

'Constants for topmost.
Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2
Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1
Public Const SWP_NOACTIVATE = &H10
Public Const SWP_SHOWWINDOW = &H40
Public Const FLAGS = SWP_NOMOVE Or SWP_NOSIZE
Public Declare Function SetWindowPos Lib "User32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Enum ONTOPSETTING
    WINDOW_ONTOP = HWND_TOPMOST
    WINDOW_NOT_ONTOP = HWND_NOTOPMOST
End Enum

Private Const GWL_EXSTYLE = (-20)
Private Const WS_EX_LAYERED = &H80000
Private Const WS_EX_TRANSPARENT = &H20&
Private Const LWA_ALPHA = &H2&
Private Declare Function FindWindow Lib "User32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function GetWindowLong Lib "User32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "User32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "User32" (ByVal hwnd As Long, ByVal crey As Byte, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Sub MakeTransparent(LhWnd As Long, bLevel As Byte)
    Dim lOldStyle As Long
    lOldStyle = GetWindowLong(LhWnd, GWL_EXSTYLE)
    SetWindowLong LhWnd, GWL_EXSTYLE, lOldStyle Or WS_EX_LAYERED
    SetLayeredWindowAttributes LhWnd, 0, bLevel, LWA_ALPHA
End Sub
'------------------------------------------------------------
' Author:  Clint M. LaFever [clint.m.lafever@cpmx.saic.com]
' Purpose:  Functionality to Set a window always on top or turn it off.
' Date: March,10 1999 @ 10:18:37
'------------------------------------------------------------
Public Sub SetFormOnTop(formHWND As Long, Optional sSETTING As ONTOPSETTING = WINDOW_ONTOP)
    On Error Resume Next
    Call SetWindowPos(formHWND, sSETTING, 0, 0, 0, 0, FLAGS)
End Sub

