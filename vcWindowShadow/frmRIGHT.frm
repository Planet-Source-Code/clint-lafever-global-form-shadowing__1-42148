VERSION 5.00
Begin VB.Form frmRIGHT 
   BackColor       =   &H00000000&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   2040
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4410
   LinkTopic       =   "Form1"
   ScaleHeight     =   2040
   ScaleWidth      =   4410
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmRIGHT"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    MakeTransparent Me.hwnd, 50
    SetFormOnTop Me.hwnd, WINDOW_ONTOP
End Sub
