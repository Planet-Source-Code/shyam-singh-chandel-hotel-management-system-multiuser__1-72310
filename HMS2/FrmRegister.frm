VERSION 5.00
Object = "{97824D60-58A5-4D7C-A442-3412CD2787CD}#2.0#0"; "USSTYLE.OCX"
Begin VB.Form FrmRegister 
   BackColor       =   &H00008080&
   BorderStyle     =   0  'None
   Caption         =   "Registration Form"
   ClientHeight    =   7485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11955
   Icon            =   "FrmRegister.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7485
   ScaleWidth      =   11955
   ShowInTaskbar   =   0   'False
   Begin Project1.USStyle USStyle8 
      Height          =   1095
      Left            =   480
      TabIndex        =   37
      Top             =   6000
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   1931
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "View Room Charges"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      PictureAlignment=   1
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
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1980
      ItemData        =   "FrmRegister.frx":0CCA
      Left            =   4800
      List            =   "FrmRegister.frx":0CCC
      TabIndex        =   35
      Top             =   5280
      Visible         =   0   'False
      Width           =   2295
   End
   Begin Project1.USStyle USStyle6 
      Height          =   495
      Left            =   4080
      TabIndex        =   34
      Top             =   6000
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Go to Confirm Registration"
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
   Begin Project1.USStyle USStyle5 
      Height          =   495
      Left            =   7680
      TabIndex        =   33
      Top             =   6600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
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
   Begin Project1.USStyle USStyle4 
      Height          =   495
      Left            =   6480
      TabIndex        =   32
      Top             =   6600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Delete"
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
   Begin Project1.USStyle USStyle3 
      Height          =   495
      Left            =   5280
      TabIndex        =   31
      Top             =   6600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Edit"
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
      Height          =   495
      Left            =   4080
      TabIndex        =   30
      Top             =   6600
      Width           =   1095
      _ExtentX        =   1931
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Add"
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
      Left            =   6240
      TabIndex        =   29
      Top             =   960
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "New Entry"
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
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFF80&
      Caption         =   "Go to Confirm Registration"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   28
      Top             =   5760
      Visible         =   0   'False
      Width           =   3135
   End
   Begin VB.TextBox Text17 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   9480
      MultiLine       =   -1  'True
      TabIndex        =   26
      Top             =   480
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H00FFFF80&
      Caption         =   "New"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   11760
      Style           =   1  'Graphical
      TabIndex        =   25
      Top             =   4320
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   5685
      ItemData        =   "FrmRegister.frx":0CCE
      Left            =   8880
      List            =   "FrmRegister.frx":0CD0
      Sorted          =   -1  'True
      TabIndex        =   23
      Top             =   1440
      Width           =   2535
   End
   Begin VB.CommandButton Command4 
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
      Height          =   375
      Left            =   3600
      Style           =   1  'Graphical
      TabIndex        =   22
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00808080&
      Caption         =   "Delete"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      Style           =   1  'Graphical
      TabIndex        =   21
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00808080&
      Caption         =   "Edit"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   20
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Add"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   360
      Style           =   1  'Graphical
      TabIndex        =   19
      Top             =   5520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7200
      TabIndex        =   7
      Top             =   4920
      Width           =   1455
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   4920
      Width           =   1935
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2640
      TabIndex        =   5
      Top             =   4920
      Width           =   2055
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   480
      TabIndex        =   4
      Top             =   4920
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   3840
      Width           =   6015
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   480
      TabIndex        =   2
      Top             =   3840
      Width           =   2055
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   1095
      Left            =   480
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   2280
      Width           =   5775
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2520
      TabIndex        =   0
      Top             =   1440
      Width           =   6135
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   480
      TabIndex        =   8
      Top             =   1440
      Width           =   1935
   End
   Begin Project1.USStyle USStyle7 
      Height          =   375
      Left            =   6720
      TabIndex        =   36
      Top             =   4920
      Width           =   375
      _ExtentX        =   661
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "V"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      Checked         =   0   'False
      ColorButtonHover=   40960
      ColorButtonUp   =   32768
      ColorButtonDown =   49152
      BorderBrightness=   0
      ColorBright     =   65280
      DisplayHand     =   0   'False
      ColorScheme     =   5
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Height          =   7440
      Left            =   15
      Top             =   15
      Width           =   11895
   End
   Begin VB.Label Label18 
      BackStyle       =   0  'Transparent
      Caption         =   "Last Customer No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7560
      TabIndex        =   27
      Top             =   480
      Width           =   2055
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   8880
      TabIndex        =   24
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Room Registration"
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
      Height          =   615
      Left            =   480
      TabIndex        =   18
      Top             =   360
      Width           =   6855
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Type of Room"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   4800
      TabIndex        =   17
      Top             =   4560
      Width           =   2175
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "No of Days"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   7200
      TabIndex        =   16
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Date of Arrival"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   15
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Registration Date"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   14
      Top             =   4560
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer No."
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   13
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2640
      TabIndex        =   12
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Telephone"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   11
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   10
      Top             =   1920
      Width           =   1575
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Name"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   1080
      Width           =   975
   End
End
Attribute VB_Name = "FrmRegister"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyDb As Database, MyRs As Recordset
Dim Db As Database, Rs As Recordset
Dim FS, SL
Const Reading = 1, Writing = 2



Private Sub Command1_Click()
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(CustomerPath & "\CUSTOMER.MDB")
Set MyRs = MyDb.OpenRecordset("customer", dbOpenDynaset)
MyRs.AddNew
     MyRs!ID = "NO"
     MyRs!SL = Text1.Text
     MyRs!Name = Text2.Text
     MyRs!ADDRESS = Text3.Text
     MyRs!TEL = Text4.Text
     MyRs!EMAIL = Text5.Text
     MyRs!REGEXPIRY = Text7.Text
     MyRs!ARRIVAL = Text7.Text
     MyRs!REGDATE = Text6.Text
     MyRs!TYPEOFROOM = Text8.Text
     MyRs!NOOFDAYS = Text9.Text
MyRs.Update
Open MainPath & "\CUSTOMERNO.TXT" For Output As #3
     Print #3, Text1.Text
     Close #3
MyDb.Close
     Text1.Text = ""
     Text2.Text = ""
     Text3.Text = ""
     Text4.Text = ""
     Text5.Text = ""
     Text7.Text = ""
     Text7.Text = ""
     Text6.Text = ""
     Text8.Text = ""
     Text9.Text = ""
     
     
Form_Load
Text2.SetFocus

End Sub

Private Sub Command2_Click()
MyRs.Edit
     MyRs!SL = Text1.Text
     MyRs!Name = Text2.Text
     MyRs!ADDRESS = Text3.Text
     MyRs!TEL = Text4.Text
     MyRs!EMAIL = Text5.Text
     MyRs!REGEXPIRY = Text7.Text
     MyRs!ARRIVAL = Text7.Text
     MyRs!REGDATE = Text6.Text
     MyRs!TYPEOFROOM = Text8.Text
     MyRs!NOOFDAYS = Text9.Text
MyRs.Update
MsgBox "Registration Record is Edited", vbOKOnly + vbInformation
MyDb.Close
     Text1.Text = ""
     Text2.Text = ""
     Text3.Text = ""
     Text4.Text = ""
     Text5.Text = ""
     Text7.Text = ""
     Text7.Text = ""
     Text6.Text = ""
     Text8.Text = ""
     Text9.Text = ""
Form_Load
Text2.SetFocus
End Sub

Private Sub Command3_Click()
res = MsgBox("Are you sure that you want to delete the Registered Customer Record.", vbYesNo + vbQuestion)
If res = vbYes Then
MyRs.Delete
Command5_Click
Form_Load
                                        'MsgBox "The Record of Registered Customer is Deleted", vbOKOnly + vbInformation
Else
Exit Sub
End If

End Sub

Private Sub Command4_Click()

Unload Me

End Sub

Private Sub Command5_Click()
     Text1.Text = ""
     Text2.Text = ""
     Text3.Text = ""
     Text4.Text = ""
     Text5.Text = ""
     Text8.Text = ""
     Text9.Text = ""
     Text1.Text = Val(Text17.Text) + 1
     Text1.Text = Format(Text1.Text, "000000")
     Text2.SetFocus
End Sub

Private Sub Command6_Click()
FrmConferm.Show
Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Left = MDIForm1.Width / 2 - Me.Width / 2 - 300
Me.Top = MDIForm1.Height / 2 - Me.Height / 2 - 500
Dim D
List1.Clear
List2.Clear
D = Format(Now, "dd/mm/yy")

Set FS = CreateObject("Scripting.FileSystemObject")
Set SL = FS.OpenTextFile(MainPath & "\CUSTOMERNO.TXT", Reading)
Text17.Text = SL.READALL
SL.Close
Text1.Text = Val(Text17.Text) + 1
Text1.Text = Format(Text1.Text, "000000")

Text6.Text = D
Text7.Text = D

SQL = "select distinct typeofroom from ROOMS"
Set Db = DBEngine.Workspaces(0).OpenDatabase(RoomsPath & "\ROOMS.MDB")
Set Rs = Db.OpenRecordset(SQL, dbOpenDynaset)
If Rs.RecordCount <= 0 Then
Exit Sub
Else
   Do While Not Rs.EOF
      List1.AddItem Rs!TYPEOFROOM
    Rs.MoveNext
    Loop
End If

'Set MyDb = DBEngine.Workspaces(0).OpenDatabase(CustomerPath & "\CUSTOMER.MDB")
'Set MyRs = MyDb.OpenRecordset("customer", dbOpenDynaset)
'If MyRs.RecordCount <= 0 Then
'Exit Sub
'Else
'   Do While Not MyRs.EOF
'      List2.AddItem MyRs!SL
'    MyRs.MoveNext
'    Loop
'End If
'MyDb.Close

SQL = "select * from CUSTOMER where ID='" & "NO" & "'"
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(CustomerPath & "\CUSTOMER.MDB")
Set MyRs = MyDb.OpenRecordset(SQL, dbOpenDynaset)

If MyRs.RecordCount <= 0 Then
Exit Sub
Else
   Do While Not MyRs.EOF
      List2.AddItem MyRs!SL
    MyRs.MoveNext
    Loop
End If

End Sub

Private Sub List1_DblClick()
Text8.Text = List1.Text
List1.Visible = False
Text9.SetFocus
End Sub

Private Sub List1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
 Text8.Text = List1.Text
 Text9.SetFocus
 List1.Visible = False
 End If
 
End Sub

Private Sub List2_Click()
SQL = "select * from CUSTOMER where SL='" & List2.Text & "'"
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(CustomerPath & "\CUSTOMER.MDB")
Set MyRs = MyDb.OpenRecordset(SQL, dbOpenDynaset)
   Text1.Text = MyRs!SL
   Text3.Text = MyRs!ADDRESS
   Text2.Text = MyRs!Name
   Text4.Text = MyRs!TEL
   Text5.Text = MyRs!EMAIL
   Text6.Text = MyRs!REGDATE
   Text7.Text = MyRs!ARRIVAL
   Text8.Text = MyRs!TYPEOFROOM
   Text9.Text = MyRs!NOOFDAYS
   
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
    'Text4.SetFocus
 End If
End Sub
Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text5.SetFocus
 End If
End Sub
Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text6.SetFocus
 End If
End Sub
Private Sub Text6_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text7.SetFocus
 End If
End Sub
Private Sub Text7_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    List1.Visible = True
    List1.SetFocus
    If List1.ListCount = 0 Then
    Exit Sub
    Else
    List1.ListIndex = 0
    End If
  End If
End Sub

Private Sub Text8_Click()
List1.Visible = True
End Sub

Private Sub Text8_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text9.SetFocus
 End If
End Sub
Private Sub Text9_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
   Command1_Click
 End If
End Sub

Private Sub USStyle1_Click()
Command5_Click
End Sub

Private Sub USStyle2_Click()
Command1_Click
End Sub

Private Sub USStyle3_Click()
Command2_Click
End Sub

Private Sub USStyle4_Click()
Command3_Click
End Sub

Private Sub USStyle5_Click()
Unload Me

End Sub

Private Sub USStyle6_Click()
Command6_Click
End Sub

Private Sub USStyle7_Click()
If List1.Visible = True Then
List1.Visible = False
Else
List1.Visible = True
End If

End Sub

Private Sub USStyle8_Click()
FrmRoomInfo.Show
End Sub
