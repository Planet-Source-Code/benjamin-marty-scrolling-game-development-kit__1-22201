VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   213
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   312
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim WithEvents TE As TileEdit
Attribute TE.VB_VarHelpID = -1

Private Sub Form_Click()
    Set TE = New TileEdit
    Set TE.Disp = New BMDXDisplay
    TE.Disp.OpenEx
    TE.Create 32, 32, 7, 5
End Sub

Sub Pause(L As Single)
    Dim T As Single
    T = Timer
    Do While Timer - T < L
        DoEvents
    Loop
End Sub

Private Sub TE_EditComplete()
    Set TE = Nothing
End Sub
