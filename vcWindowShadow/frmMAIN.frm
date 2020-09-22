VERSION 5.00
Begin VB.Form frmMAIN 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Main"
   ClientHeight    =   1140
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   1830
   Icon            =   "frmMAIN.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   1830
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSTOP 
      Caption         =   "&Stop"
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   1575
   End
   Begin VB.CommandButton cmdGO 
      Caption         =   "&Go"
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1575
   End
End
Attribute VB_Name = "frmMAIN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'------------------------------------------------------------
' For more of my stuff you can visit http://vbasic.iscool.net
'------------------------------------------------------------



Option Explicit
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Declare Function GetForegroundWindow Lib "User32" () As Long
Private Declare Function GetWindowText Lib "User32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetWindowRect Lib "User32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function MoveWindow Lib "User32" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Declare Function IsZoomed Lib "User32" (ByVal hwnd As Long) As Long
Private Declare Function IsIconic Lib "User32" (ByVal hwnd As Long) As Long
Private Declare Function IsWindowVisible Lib "User32" (ByVal hwnd As Long) As Long
Public Continue As Boolean
Public frmR As frmRIGHT, frmB As frmBOTTOM
Private Sub cmdGO_Click()
    On Error Resume Next
    MsgBox "Note, this was made real quick just because I wanted to see if something basic like this could actually work.  There are issues with the shadow forms being ontop of the taskbar and also for applications that have skinned forms with non rectangular shapes it makes them look bad, but just wanted to share this code."
    MsgBox "Also note, it will not Shadow its own main form or windows that are maximized or minimized.  Just let this run, minimize it, and go open other forms and file open windows and what not."
    Set frmR = New frmRIGHT
    Set frmB = New frmBOTTOM
    Continue = True
    Me.cmdSTOP.Enabled = True
    Me.cmdSTOP.SetFocus
    Me.cmdGO.Enabled = False
    Start
End Sub
Private Sub cmdSTOP_Click()
    On Error Resume Next
    Continue = False
    DoEvents
    Me.cmdGO.Enabled = True
    Me.cmdGO.SetFocus
    Me.cmdSTOP.Enabled = False
    Unload frmB
    Unload frmR
End Sub
Private Sub Form_Load()
    On Error Resume Next
    Continue = False
    Me.cmdSTOP.Enabled = False
End Sub
Public Sub Start()
    On Error Resume Next
    Dim h As Long, nLen As Long, nSTR As String * 255, wCAP As String, r As RECT
    While Continue = True
        h = GetForegroundWindow
        If h <> 0 Then
            nLen = GetWindowText(h, nSTR, Len(nSTR) - 1)
            If nLen <> 0 Then
                wCAP = Left(nSTR, nLen)
                If UCase(wCAP) <> "PROGRAM MANAGER" And h <> Me.hwnd And h <> frmB.hwnd And h <> frmR.hwnd Then
                    If IsZoomed(h) Or IsIconic(h) Or IsWindowVisible(h) = False Then
                        frmB.Visible = False
                        frmR.Visible = False
                    Else
                        GetWindowRect h, r
                        MoveWindow frmB.hwnd, r.Left + 10, r.Bottom, r.Right - r.Left, 10, True
                        MoveWindow frmR.hwnd, r.Right, r.Top + 10, 10, r.Bottom - r.Top - 10, True
                        If frmB.Visible = False Then
                            frmB.Visible = True
                            SetFormOnTop frmB.hwnd, WINDOW_ONTOP
                        End If
                        If frmR.Visible = False Then
                            frmR.Visible = True
                            SetFormOnTop frmR.hwnd, WINDOW_ONTOP
                        End If
                    End If
                Else
                    frmB.Visible = False
                    frmR.Visible = False
                End If
            Else
                frmB.Visible = False
                frmR.Visible = False
            End If
        End If
        Sleep 1: DoEvents
    Wend
End Sub
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    On Error Resume Next
    Dim frm As Form
    If Me.cmdSTOP.Enabled = True Then
        cmdSTOP_Click
        Cancel = True
        MsgBox "Shadowing stopped.  Click again to close."
    Else
        Continue = False
        DoEvents
        For Each frm In Forms
            If frm.Name < Me.Name Then Unload frm
        Next
        Unload Me
    End If
End Sub
