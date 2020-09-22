VERSION 5.00
Object = "{97824D60-58A5-4D7C-A442-3412CD2787CD}#2.0#0"; "USSTYLE.OCX"
Begin VB.Form FrmLogin 
   BackColor       =   &H00008080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Login"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6045
   Icon            =   "FrmLogin.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   6045
   Begin Project1.USStyle USStyle4 
      Height          =   375
      Left            =   4440
      TabIndex        =   11
      Top             =   120
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Help ?"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   4210752
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   15309136
      ColorButtonUp   =   13657888
      ColorButtonDown =   10512144
      BorderBrightness=   0
      ColorBright     =   16772528
      DisplayHand     =   0   'False
      ColorScheme     =   1
   End
   Begin Project1.USStyle USStyle3 
      Height          =   375
      Left            =   3600
      TabIndex        =   10
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Login"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   65535
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   40960
      ColorButtonUp   =   32768
      ColorButtonDown =   49152
      BorderBrightness=   0
      ColorBright     =   65280
      DisplayHand     =   0   'False
      ColorScheme     =   5
   End
   Begin Project1.USStyle USStyle2 
      Height          =   375
      Left            =   3600
      TabIndex        =   9
      Top             =   1320
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Quit"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   65535
      Style           =   7
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
      Height          =   375
      Left            =   3600
      TabIndex        =   8
      Top             =   720
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Login"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   65535
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   40960
      ColorButtonUp   =   32768
      ColorButtonDown =   49152
      BorderBrightness=   0
      ColorBright     =   65280
      DisplayHand     =   0   'False
      ColorScheme     =   5
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00FFFF00&
      Caption         =   "Help ?"
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
      Left            =   7920
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   120
      Width           =   1455
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00808080&
      Caption         =   "Login"
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
      Left            =   8520
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   720
      Width           =   1455
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1320
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Login"
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
      Left            =   8040
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   720
      Width           =   1455
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      IMEMode         =   3  'DISABLE
      Left            =   1800
      PasswordChar    =   "."
      TabIndex        =   1
      Top             =   1320
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   1800
      TabIndex        =   0
      Top             =   720
      Width           =   1695
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Password"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   3
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "User Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   840
      Width           =   1455
   End
End
Attribute VB_Name = "FrmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyDb As Database, MyRs As Recordset
Public logdb

Private Sub Command1_Click()
On Error Resume Next
Dim USER, PASS

SQL = "SELECT * FROM USERDB WHERE UserName='" & Text1.Text & "'"
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(UserPath & "\USER.MDB")
Set MyRs = MyDb.OpenRecordset(SQL, dbOpenDynaset)
USER = MyRs!UserName
PASS = MyRs!PASS

If Text1.Text = "" Then
MsgBox "Please Enter the User Name."
Text1.SetFocus
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "Please Enter the Password."
Text2.SetFocus
Exit Sub
End If

If Text1.Text = USER And Text2.Text = PASS Then
MDIForm1.Show
Unload FrmSplash
Unload Me
PlayWave App.Path & "\123\DRUM CLICK SOUND.wav"
Exit Sub
Else

MsgBox "WRONG USER NAME AND PASSWORD"
PlayWave App.Path & "\123\ERROR LAFING.wav"
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
End If


MyDb.Close

End Sub

Private Sub Command2_Click()
End
End Sub

Private Sub Command3_Click()
On Error Resume Next

If Text1.Text = "" Then
MsgBox "Please Enter the User Name."
Text1.SetFocus
Exit Sub
End If
If Text2.Text = "" Then
MsgBox "Please Enter the Password."
Text2.SetFocus
Exit Sub
End If

If Text1.Text = "rkm" And Text2.Text = "rkm" Then
MDIForm1.Show
Unload FrmSplash
Unload Me
PlayWave App.Path & "\123\DRUM CLICK SOUND.wav"
Else
MsgBox "Wrong User Name or Password"
PlayWave App.Path & "\123\ERROR LAFING.wav"
Text1.Text = ""
Text2.Text = ""
Text1.SetFocus
Exit Sub
End If

End Sub

Private Sub Command4_Click()
FrmHelp2.Show
End Sub

Private Sub Form_Load()
Me.Top = FrmSplash.Height - Me.Height - 1000
Me.Left = FrmSplash.Left + 1600

Set MyDb = DBEngine.Workspaces(0).OpenDatabase(UserPath & "\USER.MDB")
Set MyRs = MyDb.OpenRecordset("USERDB", dbOpenDynaset)

If MyRs.RecordCount <= 0 Then
Command3.Visible = True
USStyle3.Visible = True
logdb = "No"
Exit Sub
Else
Command3.Visible = False
USStyle3.Visible = False
logdb = "Yes"

End If
 MyRs.Close
 MyDb.Close
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   Text2.SetFocus
   End If
   
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
 If logdb = "Yes" Then
 Command1_Click
 Else
  Command3_Click
 End If
 End If
 
End Sub

Private Sub USStyle1_Click()
Command1_Click
End Sub

Private Sub USStyle2_Click()
Unload Me
Command2_Click
End Sub

Private Sub USStyle3_Click()
Command3_Click
End Sub

Private Sub USStyle4_Click()
FrmHelp2.Show
End Sub
