VERSION 5.00
Object = "{97824D60-58A5-4D7C-A442-3412CD2787CD}#2.0#0"; "USSTYLE.OCX"
Begin VB.Form FrmRoomInfo 
   BackColor       =   &H00008080&
   BorderStyle     =   0  'None
   Caption         =   "Room Information"
   ClientHeight    =   4965
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   ScaleHeight     =   4965
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "Quit"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7660
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5160
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Update Room"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7660
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   3720
      Width           =   3255
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Height          =   3345
      Left            =   585
      Sorted          =   -1  'True
      TabIndex        =   6
      Top             =   1200
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3820
      TabIndex        =   5
      Top             =   1800
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   3820
      TabIndex        =   4
      Top             =   1200
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3820
      TabIndex        =   3
      Top             =   2400
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   1005
      Left            =   3820
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00808080&
      Caption         =   "Edit Room Record"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7660
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   4200
      Width           =   3255
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H00808080&
      Caption         =   "Delete Room Record"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7660
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   4680
      Width           =   3255
   End
   Begin Project1.USStyle USStyle4 
      Height          =   495
      Left            =   2880
      TabIndex        =   9
      Top             =   4080
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
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
      ForeColor       =   4210752
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
   Begin Project1.USStyle USStyle3 
      Height          =   495
      Left            =   8880
      TabIndex        =   10
      Top             =   3360
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   873
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Delete Room"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   4210752
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
      Height          =   495
      Left            =   8760
      TabIndex        =   11
      Top             =   3000
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   873
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Edit Room Record"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   4210752
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
      Height          =   495
      Left            =   8760
      TabIndex        =   12
      Top             =   2520
      Width           =   3255
      _ExtentX        =   5741
      _ExtentY        =   873
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Update Room"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   4210752
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Room No."
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
      Index           =   1
      Left            =   580
      TabIndex        =   18
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Room Charges"
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
      Left            =   2500
      TabIndex        =   17
      Top             =   1800
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Room No."
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
      Index           =   0
      Left            =   2500
      TabIndex        =   16
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Features"
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
      Left            =   2500
      TabIndex        =   15
      Top             =   3000
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Type of Room"
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
      Index           =   2
      Left            =   2500
      TabIndex        =   14
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "View Room Charges"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   495
      Index           =   3
      Left            =   100
      TabIndex        =   13
      Top             =   360
      Width           =   6735
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      Height          =   4935
      Left            =   10
      Top             =   10
      Width           =   6975
   End
End
Attribute VB_Name = "FrmRoomInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyDb As Database, MyRs As Recordset

Private Sub Command1_Click()
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(RoomsPath & "\ROOMS.MDB")
Set MyRs = MyDb.OpenRecordset("ROOMS", dbOpenDynaset)
MyRs.AddNew
    MyRs!ROOMNO = Text1.Text
    MyRs!TYPEOFROOM = Text3.Text
    MyRs!Rate = Text2.Text
    MyRs!FEATURS = Text4.Text
    MyRs!Status = "EMPTY"
 MyRs.Update
 MsgBox "RECORD OF ROOM IS UPDATED", vbOKOnly, "Hotel Management System"
 Text1 = ""
 Text2 = ""
 Text3 = ""
 Text4 = ""
 Form_Load
 
End Sub

Private Sub Command2_Click()
 
 Unload Me
 
End Sub

Private Sub Command3_Click()

MyRs.Edit
    MyRs!ROOMNO = Text1.Text
    MyRs!TYPEOFROOM = Text3.Text
    MyRs!Rate = Text2.Text
    MyRs!FEATURS = Text4.Text
 MyRs.Update
 MyRs.Close
 MyDb.Close
 MsgBox "RECORD OF ROOM IS UPDATED", vbOKOnly, "Hotel Management System"
 Text1 = ""
 Text2 = ""
 Text3 = ""
 Text4 = ""
 Command3.Enabled = False
 Form_Load
End Sub

Private Sub Command4_Click()
ans = MsgBox("Do you want to Delete Room No. " & List1.Text, vbQuestion + vbYesNo, "Hotel Management System")
If ans = vbYes Then
MyRs.Delete
MsgBox "Room No. " & List1.Text & " has been deleted.", vbOKOnly, "Hotel Management System"
Text1 = ""
 Text2 = ""
 Text3 = ""
 Text4 = ""
 Form_Load
 Else
 Exit Sub
 End If
 
End Sub

Private Sub Form_Load()
Me.Left = MDIForm1.Width / 2 - Me.Width / 2 - 300 '4000
Me.Top = MDIForm1.Height / 2 - Me.Height / 2 - 500


List1.Clear

Set MyDb = DBEngine.Workspaces(0).OpenDatabase(RoomsPath & "\ROOMS.MDB")
Set MyRs = MyDb.OpenRecordset("ROOMS", dbOpenDynaset)
If MyRs.RecordCount <= 0 Then
Exit Sub
Else
   Do While Not MyRs.EOF
      List1.AddItem MyRs!ROOMNO
    MyRs.MoveNext
    Loop
End If

MyDb.Close

End Sub

Private Sub List1_Click()
Command3.Enabled = True
Command4.Enabled = True
 USStyle2.Enabled = True
 USStyle3.Enabled = True
SQL = "select * from rooms where roomno='" & List1.Text & "'"

Set MyDb = DBEngine.Workspaces(0).OpenDatabase(RoomsPath & "\ROOMS.MDB")
'Set MyRs = MyDb.OpenRecordset("ROOMS", dbOpenDynaset)
Set MyRs = MyDb.OpenRecordset(SQL, dbOpenDynaset)

   Text1.Text = MyRs!ROOMNO
   Text3.Text = MyRs!TYPEOFROOM
   Text2.Text = MyRs!Rate
   Text4.Text = MyRs!FEATURS
Label1(3).Caption = "Room No. " & List1.Text & " with following information"

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  Text2.SetFocus
End If
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  Text3.SetFocus
End If
End Sub
Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  Text4.SetFocus
  USStyle1.Enabled = True
  Command1.Enabled = True
End If
End Sub

Private Sub Text4_Click()
Command1.Enabled = True
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  Command1.Enabled = True
  USStyle1_Click
  'Command1_Click
  Command1.Enabled = False
  Text1.SetFocus
End If
End Sub

Private Sub USStyle1_Click()
Command1_Click
End Sub

Private Sub USStyle4_Click()
Unload Me

End Sub

