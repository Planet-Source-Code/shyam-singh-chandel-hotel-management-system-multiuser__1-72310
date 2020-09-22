VERSION 5.00
Object = "{97824D60-58A5-4D7C-A442-3412CD2787CD}#2.0#0"; "USSTYLE.OCX"
Begin VB.Form FrmCreatUser 
   BackColor       =   &H00008080&
   BorderStyle     =   0  'None
   Caption         =   "Creat User"
   ClientHeight    =   4455
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6405
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   4455
   ScaleWidth      =   6405
   ShowInTaskbar   =   0   'False
   Begin Project1.USStyle USStyle2 
      Height          =   615
      Left            =   2400
      TabIndex        =   9
      Top             =   3480
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Exit"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   6
      Checked         =   0   'False
      ColorButtonHover=   40960
      ColorButtonUp   =   32768
      ColorButtonDown =   49152
      BorderBrightness=   0
      ColorBright     =   65280
      DisplayHand     =   0   'False
      ColorScheme     =   5
   End
   Begin Project1.USStyle USStyle1 
      Height          =   615
      Left            =   2400
      TabIndex        =   8
      Top             =   2760
      Width           =   3375
      _ExtentX        =   5953
      _ExtentY        =   1085
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Creat User"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Style           =   6
      Checked         =   0   'False
      ColorButtonHover=   40960
      ColorButtonUp   =   32768
      ColorButtonDown =   49152
      BorderBrightness=   0
      ColorBright     =   65280
      DisplayHand     =   0   'False
      ColorScheme     =   5
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3120
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Create User"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2400
      Visible         =   0   'False
      Width           =   3255
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Height          =   2955
      Left            =   480
      Sorted          =   -1  'True
      TabIndex        =   4
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   1920
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   3840
      TabIndex        =   0
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Height          =   4455
      Left            =   0
      Top             =   0
      Width           =   6375
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Creat User"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   735
      Left            =   360
      TabIndex        =   10
      Top             =   240
      Width           =   5655
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Index           =   1
      Left            =   480
      TabIndex        =   5
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   495
      Left            =   2400
      TabIndex        =   3
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000040&
      Height          =   375
      Index           =   0
      Left            =   2400
      TabIndex        =   2
      Top             =   1200
      Width           =   975
   End
End
Attribute VB_Name = "FrmCreatUser"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyDb As Database, MyRs As Recordset

Private Sub Command1_Click()
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(UserPath & "\USER.MDB")
Set MyRs = MyDb.OpenRecordset("USERDB", dbOpenDynaset)
If Text1.Text = "" And Text2.Text = "" Then
MsgBox "Plaes enter the user name and password"
Exit Sub
End If

MyRs.AddNew
     MyRs!UserName = Text1.Text
     MyRs!PASS = Text2.Text
MyRs.Update
MsgBox "User has been created"

MyDb.Close
Text1 = ""
Text2 = ""
Form_Load
     
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()

Me.Left = MDIForm1.Width / 2 - Me.Width / 2 - 300
Me.Top = MDIForm1.Height / 2 - Me.Height / 2 - 500

List1.Clear
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(UserPath & "\USER.MDB")
Set MyRs = MyDb.OpenRecordset("USERDB", dbOpenDynaset)
If MyRs.RecordCount = 0 Then
Exit Sub
Else
Do While Not MyRs.EOF
List1.AddItem MyRs!UserName
MyRs.MoveNext
Loop
MyDb.Close

End If

End Sub

Private Sub USStyle1_Click()
Command1_Click
End Sub

Private Sub USStyle2_Click()
Unload Me
End Sub
