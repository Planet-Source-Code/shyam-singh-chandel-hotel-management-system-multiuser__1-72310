VERSION 5.00
Begin VB.Form FrmSettings 
   BackColor       =   &H00008080&
   Caption         =   "Data Storage Path Settings"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   5760
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3720
      TabIndex        =   2
      Top             =   1920
      Width           =   1335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Apply"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2280
      TabIndex        =   1
      Top             =   1920
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   600
      TabIndex        =   0
      Top             =   960
      Width           =   4455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Data Storage Path:  E. g.  ""D:\HotelData"""
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   600
      TabIndex        =   3
      Top             =   720
      Width           =   4455
   End
End
Attribute VB_Name = "FrmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
 
 On Error Resume Next
 Call SaveSetting("USHMS", "SETTINGS", "0US001", Text1.Text)
 Call SaveSetting("USHMS", "SETTINGS", "0US002", Text1.Text & "\Customer")
 Call SaveSetting("USHMS", "SETTINGS", "0US003", Text1.Text & "\PRINT")
 Call SaveSetting("USHMS", "SETTINGS", "0US004", Text1.Text & "\STAFF")
 Call SaveSetting("USHMS", "SETTINGS", "0US005", Text1.Text & "\ITEMS")
 Call SaveSetting("USHMS", "SETTINGS", "0US006", Text1.Text & "\USER")
 Call SaveSetting("USHMS", "SETTINGS", "0US007", Text1.Text & "\ROOMS")
 Call SaveSetting("USHMS", "SETTINGS", "0US008", Text1.Text & "\RESTO")
 
 MkDir App.Path & "\SCR"
 
 FileCopy App.Path & "\1.JPG", App.Path & "\SCR\1.JPG"
 FileCopy App.Path & "\2.JPG", App.Path & "\SCR\2.JPG"
 FileCopy App.Path & "\3.JPG", App.Path & "\SCR\3.JPG"
 FileCopy App.Path & "\4.JPG", App.Path & "\SCR\4.JPG"
 FileCopy App.Path & "\5.JPG", App.Path & "\SCR\5.JPG"
 FileCopy App.Path & "\6.JPG", App.Path & "\SCR\6.JPG"
 
 Kill App.Path & "\1.JPG"
 Kill App.Path & "\2.JPG"
 Kill App.Path & "\3.JPG"
 Kill App.Path & "\4.JPG"
 Kill App.Path & "\5.JPG"
 Kill App.Path & "\6.JPG"
 
 FrmSplash.Show
 
 Unload Me
 
End Sub
                                                        
Private Sub Command2_Click()
End
End Sub
