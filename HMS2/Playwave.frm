VERSION 5.00
Begin VB.Form FrmPlayWave 
   Caption         =   "PlayWaveFiles"
   ClientHeight    =   2055
   ClientLeft      =   60
   ClientTop       =   330
   ClientWidth     =   4365
   LinkTopic       =   "Form1"
   ScaleHeight     =   2055
   ScaleWidth      =   4365
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Play the MONTHS.WAV file"
      Height          =   495
      Left            =   360
      TabIndex        =   1
      Top             =   960
      Width           =   2655
   End
   Begin VB.CommandButton Sound 
      Caption         =   "Play the DAYS.WAV file"
      Height          =   495
      Left            =   360
      TabIndex        =   0
      Top             =   240
      Width           =   2655
   End
End
Attribute VB_Name = "FrmPlayWave"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Sound_Click()
PlayWave App.Path & "\123\ERROR LAFING.wav"
End Sub
