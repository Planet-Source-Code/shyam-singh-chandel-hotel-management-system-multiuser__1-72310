VERSION 5.00
Begin VB.Form FrmSplash 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   8730
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   9585
   Icon            =   "FrmSplash.frx":0000
   LinkTopic       =   "Form1"
   Picture         =   "FrmSplash.frx":0CCA
   ScaleHeight     =   8730
   ScaleWidth      =   9585
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer2 
      Interval        =   500
      Left            =   600
      Top             =   3000
   End
   Begin VB.Timer Timer1 
      Interval        =   3000
      Left            =   360
      Top             =   1800
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Shyam Singh Chandel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   975
      Left            =   3700
      TabIndex        =   5
      Top             =   570
      Width           =   5775
   End
   Begin VB.Label lblCompany 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Developed by"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   240
      Width           =   7095
   End
   Begin VB.Label lblLicenseTo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Shyam Singh Chandel"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   3720
      TabIndex        =   3
      Top             =   600
      Width           =   5775
   End
   Begin VB.Image Image2 
      Height          =   3000
      Left            =   0
      Picture         =   "FrmSplash.frx":2C026
      Top             =   -720
      Width           =   3750
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H80000006&
      BorderWidth     =   3
      FillColor       =   &H00FFFFFF&
      Height          =   8730
      Left            =   15
      Top             =   15
      Width           =   9540
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Welcome to"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080C0FF&
      Height          =   375
      Left            =   9120
      TabIndex        =   2
      Top             =   1440
      Width           =   5655
   End
   Begin VB.Image Image1 
      Height          =   2775
      Left            =   0
      Picture         =   "FrmSplash.frx":2D634
      Stretch         =   -1  'True
      Top             =   6000
      Width           =   9690
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HOTEL MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   855
      Left            =   9480
      TabIndex        =   1
      Top             =   3240
      Width           =   5655
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "HOTEL MANAGEMENT SYSTEM"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   855
      Left            =   9240
      TabIndex        =   0
      Top             =   3360
      Width           =   5655
   End
End
Attribute VB_Name = "FrmSplash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Function fCreateShellLink Lib "Vb5stkit.dll" (ByVal lpstrFolderName As String, ByVal lpstrLinkName As String, ByVal lpstrLinkPath As String, ByVal lpstrLinkArgs As String) As Long

Private Sub Form_Load()
Dim lReturn As Long
  On Error Resume Next
   lReturn = fCreateShellLink("..\..\Desktop", "USHMS", App.Path & "\HotelManagement.exe", "")
  
  MainPath = GetSetting("USHMS", "SETTINGS", "0US001")
  CustomerPath = GetSetting("USHMS", "SETTINGS", "0US002")
  PrintPath = GetSetting("USHMS", "SETTINGS", "0US003")
  StaffPath = GetSetting("USHMS", "SETTINGS", "0US004")
  ItemsPath = GetSetting("USHMS", "SETTINGS", "0US005")
  UserPath = GetSetting("USHMS", "SETTINGS", "0US006")
  RoomsPath = GetSetting("USHMS", "SETTINGS", "0US007")
  RestoPath = GetSetting("USHMS", "SETTINGS", "0US008")
  SCRTIME = GetSetting("USHMS", "SETTINGS", "0US009")
  If MainPath = "" Then
  FrmSettings.Show
  Unload Me
  Exit Sub
  End If
  
  MkDir MainPath
  MkDir CustomerPath
  MkDir PrintPath
  MkDir StaffPath
  MkDir ItemsPath
  MkDir UserPath
  MkDir RoomsPath
  MkDir RestoPath
 Call CUSTOMER(CustomerPath & "\customer.mdb")
 Call STAFF(StaffPath & "\STAFF.mdb")
 Call PRINTDB(PrintPath & "\PRINT.mdb")
 Call ITEMS(ItemsPath & "\ITEMS.mdb")
 Call USERDB(UserPath & "\USER.mdb")
 Call ROOMS(RoomsPath & "\ROOMS.mdb")
 Call RESTO(RestoPath & "\RESTO.mdb")

    
End Sub

Private Sub Timer1_Timer()
FrmLogin.Show
'Me.Left = 1
'Me.Top = 1

End Sub

Private Sub Timer2_Timer()
 PlayWave App.Path & "\123\windows start.wav"
 Timer2.Enabled = False
End Sub
