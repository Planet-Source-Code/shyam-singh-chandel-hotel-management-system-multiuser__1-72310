VERSION 5.00
Object = "{97824D60-58A5-4D7C-A442-3412CD2787CD}#2.0#0"; "USSTYLE.OCX"
Begin VB.Form FrmGoodBy 
   BackColor       =   &H00FFFF00&
   BorderStyle     =   0  'None
   ClientHeight    =   1140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6075
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1140
   ScaleWidth      =   6075
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin Project1.USStyle USStyle5 
      Height          =   1335
      Left            =   -120
      TabIndex        =   0
      Top             =   -120
      Width           =   6255
      _ExtentX        =   11033
      _ExtentY        =   2355
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   48
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Good Bye...."
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   16711680
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16760976
      ColorButtonUp   =   15309136
      ColorButtonDown =   15309136
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   240
      Top             =   2760
   End
End
Attribute VB_Name = "FrmGoodBy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
MDIForm1.Visible = False
End Sub

Private Sub Timer1_Timer()
PlayWave App.Path & "\123\windows exit.wav"
End
Timer1.Enabled = False
End Sub
