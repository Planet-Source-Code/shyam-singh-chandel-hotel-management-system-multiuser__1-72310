VERSION 5.00
Object = "{97824D60-58A5-4D7C-A442-3412CD2787CD}#2.0#0"; "USSTYLE.OCX"
Begin VB.Form FrmRoomStatus 
   BackColor       =   &H00008080&
   BorderStyle     =   0  'None
   Caption         =   "Room Status"
   ClientHeight    =   8220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   10485
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8220
   ScaleWidth      =   10485
   ShowInTaskbar   =   0   'False
   Begin Project1.USStyle USStyle1 
      Height          =   495
      Left            =   6960
      TabIndex        =   10
      Top             =   7320
      Width           =   2895
      _ExtentX        =   5106
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
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
   Begin VB.OptionButton Option3 
      BackColor       =   &H00008080&
      Caption         =   "All Rooms"
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
      Left            =   8040
      TabIndex        =   9
      Top             =   480
      Value           =   -1  'True
      Width           =   1935
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00008080&
      Caption         =   "Full Rooms"
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
      Left            =   5880
      TabIndex        =   8
      Top             =   480
      Width           =   1935
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00008080&
      Caption         =   "Empty Rooms"
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
      Left            =   3600
      TabIndex        =   7
      Top             =   480
      Width           =   1935
   End
   Begin VB.ListBox List4 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   5685
      ItemData        =   "FrmPrintRestBill.frx":0000
      Left            =   6840
      List            =   "FrmPrintRestBill.frx":0002
      TabIndex        =   6
      Top             =   1440
      Width           =   3015
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   5685
      ItemData        =   "FrmPrintRestBill.frx":0004
      Left            =   3840
      List            =   "FrmPrintRestBill.frx":0006
      TabIndex        =   5
      Top             =   1440
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
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
      Height          =   375
      Left            =   600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   8520
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   5685
      ItemData        =   "FrmPrintRestBill.frx":0008
      Left            =   1920
      List            =   "FrmPrintRestBill.frx":000A
      TabIndex        =   1
      Top             =   1440
      Width           =   1935
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   5685
      ItemData        =   "FrmPrintRestBill.frx":000C
      Left            =   600
      List            =   "FrmPrintRestBill.frx":000E
      TabIndex        =   0
      Top             =   1440
      Width           =   1335
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Height          =   8175
      Left            =   0
      Top             =   0
      Width           =   10455
   End
   Begin VB.Line Line3 
      X1              =   6840
      X2              =   6840
      Y1              =   1440
      Y2              =   1080
   End
   Begin VB.Line Line2 
      X1              =   3840
      X2              =   3840
      Y1              =   1440
      Y2              =   1080
   End
   Begin VB.Line Line1 
      X1              =   1920
      X2              =   1920
      Y1              =   1440
      Y2              =   1080
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Room Status"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   615
      Left            =   600
      TabIndex        =   4
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Room No.   Room Status           Type of Room                     Feautres"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   600
      TabIndex        =   3
      Top             =   1080
      Width           =   9255
   End
End
Attribute VB_Name = "FrmRoomStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_Load()
Me.Left = MDIForm1.Width / 2 - Me.Width / 2
Me.Top = MDIForm1.Height / 2 - Me.Height / 2 - 500
List1.Clear
List2.Clear
List3.Clear
List4.Clear
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(RoomsPath & "\ROOMS.MDB")
Set MyRs = MyDb.OpenRecordset("ROOMS", dbOpenDynaset)
If MyRs.RecordCount <= 0 Then
Exit Sub
Else
   Do While Not MyRs.EOF
      List1.AddItem MyRs!ROOMNO
      List2.AddItem MyRs!Status
      List3.AddItem MyRs!TYPEOFROOM
      List4.AddItem MyRs!FEATURS
    MyRs.MoveNext
    Loop
End If
MyDb.Close

End Sub

Private Sub List1_Click()
List2.ListIndex = List1.ListIndex
List3.ListIndex = List1.ListIndex
List4.ListIndex = List1.ListIndex

End Sub

Private Sub List1_DblClick()
FrmRoomInfo.List1.Text = List1.Text
FrmRoomInfo.Show

End Sub

Private Sub List2_Click()
List1.ListIndex = List2.ListIndex
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
List1.Clear
List2.Clear
List3.Clear
List4.Clear
SQL = "select * from ROOMS where STATUS='" & "EMPTY" & "'"
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(RoomsPath & "\ROOMS.MDB")
Set MyRs = MyDb.OpenRecordset(SQL, dbOpenDynaset)
If MyRs.RecordCount <= 0 Then
Exit Sub
Else
   Do While Not MyRs.EOF
      List1.AddItem MyRs!ROOMNO
      List2.AddItem MyRs!Status
      List3.AddItem MyRs!TYPEOFROOM
      List4.AddItem MyRs!FEATURS
    MyRs.MoveNext
    Loop
End If
MyDb.Close
Else
List1.Clear
List2.Clear
List3.Clear
List4.Clear
Form_Load
Exit Sub
End If

End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
List1.Clear
List2.Clear
List3.Clear
List4.Clear
SQL = "select * from ROOMS where STATUS='" & "OCCUPIED" & "'"
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(RoomsPath & "\ROOMS.MDB")
Set MyRs = MyDb.OpenRecordset(SQL, dbOpenDynaset)
If MyRs.RecordCount <= 0 Then
Exit Sub
Else
   Do While Not MyRs.EOF
      List1.AddItem MyRs!ROOMNO
      List2.AddItem MyRs!Status
      List3.AddItem MyRs!TYPEOFROOM
      List4.AddItem MyRs!FEATURS
    MyRs.MoveNext
    Loop
End If
MyDb.Close
Else
List1.Clear
List2.Clear
List3.Clear
List4.Clear
Form_Load
Exit Sub
End If

End Sub

Private Sub Option3_Click()
If Option3.Value = True Then
Form_Load
End If
End Sub

Private Sub USStyle1_Click()
Unload Me

End Sub
