VERSION 5.00
Object = "{97824D60-58A5-4D7C-A442-3412CD2787CD}#2.0#0"; "USSTYLE.OCX"
Begin VB.Form frmAbout 
   BackColor       =   &H00008080&
   BorderStyle     =   0  'None
   ClientHeight    =   8070
   ClientLeft      =   210
   ClientTop       =   1365
   ClientWidth     =   12165
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmAbout.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8070
   ScaleWidth      =   12165
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      BackColor       =   &H0000C000&
      Height          =   6930
      Left            =   240
      TabIndex        =   3
      Top             =   240
      Width           =   11640
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Contect us: -"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   375
         Left            =   2400
         TabIndex        =   14
         Top             =   3480
         Width           =   3375
      End
      Begin VB.Image Image1 
         Height          =   3000
         Left            =   -120
         Picture         =   "frmAbout.frx":0CCA
         Top             =   -240
         Width           =   3750
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "System Requirment: -"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3720
         TabIndex        =   12
         Top             =   1560
         Width           =   2985
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "400 MHz and above Processor,  64 MB RAM and above   MSAccess 2000 and above"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1275
         Left            =   6840
         TabIndex        =   11
         Top             =   1560
         Width           =   4770
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Hotel Management System"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   27.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   660
         Left            =   3720
         TabIndex        =   10
         Top             =   600
         Width           =   7110
      End
      Begin VB.Label lblLicenseTo 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Shyam Singh Chandel"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   20.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00404040&
         Height          =   975
         Left            =   4680
         TabIndex        =   9
         Top             =   5760
         Width           =   7095
      End
      Begin VB.Label lblPlatform 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Platform: -"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   3720
         TabIndex        =   8
         Top             =   1200
         Width           =   1440
      End
      Begin VB.Label lblWarning 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Email: - shyamschandel@rediffmail.com"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   435
         Left            =   3000
         TabIndex        =   7
         Top             =   3840
         Width           =   6975
      End
      Begin VB.Label lblCompany 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Developed by"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   21.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFF00&
         Height          =   495
         Left            =   4680
         TabIndex        =   6
         Top             =   5280
         Width           =   7095
      End
      Begin VB.Label lblProductName 
         AutoSize        =   -1  'True
         Caption         =   "Product"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   32.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   11640
         TabIndex        =   5
         Top             =   1080
         Visible         =   0   'False
         Width           =   2430
      End
      Begin VB.Label lblCompanyProduct 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Win98, Win2000, WinME, XP"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5280
         TabIndex        =   4
         Top             =   1200
         Width           =   3810
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   2000
      Left            =   7200
      Top             =   2400
   End
   Begin Project1.USStyle USStyle1 
      Height          =   495
      Left            =   9960
      TabIndex        =   13
      Top             =   7320
      Width           =   1935
      _ExtentX        =   3413
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Exit"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   65535
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   16744576
      ColorButtonUp   =   16711680
      ColorButtonDown =   16761024
      BorderBrightness=   0
      ColorBright     =   16711680
      DisplayHand     =   0   'False
      ColorScheme     =   0
   End
   Begin VB.Shape Shape4 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Height          =   8055
      Left            =   0
      Top             =   0
      Width           =   12135
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "3.  Shmpii"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   2040
      TabIndex        =   2
      Top             =   3480
      Width           =   4695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "2.  Rakhi"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   3000
      Width           =   4695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "1.  Shyam Singh Chandel"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2040
      TabIndex        =   0
      Top             =   2520
      Width           =   4695
   End
   Begin VB.Shape Shape3 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   5
      Height          =   3855
      Left            =   945
      Top             =   810
      Width           =   6825
   End
   Begin VB.Shape Shape2 
      BorderColor     =   &H000000FF&
      BorderWidth     =   5
      Height          =   4035
      Left            =   840
      Top             =   720
      Width           =   7035
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00FF0000&
      BorderWidth     =   5
      Height          =   4215
      Left            =   750
      Top             =   630
      Width           =   7215
   End
   Begin VB.Image imgLogo 
      Height          =   6000
      Left            =   600
      Picture         =   "frmAbout.frx":22D8
      Top             =   840
      Visible         =   0   'False
      Width           =   7500
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
Me.Left = MDIForm1.Width / 2 - Me.Width / 2 - 300
Me.Top = MDIForm1.Height / 2 - Me.Height / 2 - 500

End Sub

Private Sub Frame1_Click()
    Unload Me
End Sub

Private Sub imgLogo_Click()
Unload Me

End Sub

Private Sub Label1_Click()
Unload Me
End Sub

Private Sub Label2_Click()
Unload Me
End Sub

Private Sub Label3_Click()
Unload Me
End Sub


Private Sub USStyle1_Click()
Unload Me
PlayWave App.Path & "\123\DRUM CLICK SOUND.wav"
End Sub
