VERSION 5.00
Begin VB.Form FrmIdle 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   345
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   330
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   345
   ScaleWidth      =   330
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "FrmIdle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.Left = 15000
    mIDLE.IDLE = SCRTIME '1 second
    mIDLE.Init Me.hwnd, 10
    Me.Visible = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
mIDLE.Terminate
End Sub
