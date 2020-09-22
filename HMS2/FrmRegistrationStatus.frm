VERSION 5.00
Object = "{97824D60-58A5-4D7C-A442-3412CD2787CD}#2.0#0"; "USSTYLE.OCX"
Begin VB.Form FrmRegistrationStatus 
   BackColor       =   &H00008080&
   BorderStyle     =   0  'None
   Caption         =   "Registration Status"
   ClientHeight    =   7845
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11325
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7845
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   Begin Project1.USStyle USStyle1 
      Height          =   495
      Left            =   8280
      TabIndex        =   8
      Top             =   6960
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
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
   Begin VB.ListBox List5 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   5490
      ItemData        =   "FrmRegistrationStatus.frx":0000
      Left            =   7680
      List            =   "FrmRegistrationStatus.frx":0002
      TabIndex        =   7
      Top             =   1320
      Width           =   3135
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   5490
      ItemData        =   "FrmRegistrationStatus.frx":0004
      Left            =   360
      List            =   "FrmRegistrationStatus.frx":0006
      TabIndex        =   4
      Top             =   1320
      Width           =   1575
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   5490
      ItemData        =   "FrmRegistrationStatus.frx":0008
      Left            =   1920
      List            =   "FrmRegistrationStatus.frx":000A
      TabIndex        =   3
      Top             =   1320
      Width           =   1695
   End
   Begin VB.CommandButton Command1 
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
      Left            =   8640
      TabIndex        =   2
      Top             =   8880
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   5490
      ItemData        =   "FrmRegistrationStatus.frx":000C
      Left            =   3600
      List            =   "FrmRegistrationStatus.frx":000E
      TabIndex        =   1
      Top             =   1320
      Width           =   2055
   End
   Begin VB.ListBox List4 
      Appearance      =   0  'Flat
      BackColor       =   &H0080C0FF&
      Height          =   5490
      ItemData        =   "FrmRegistrationStatus.frx":0010
      Left            =   5640
      List            =   "FrmRegistrationStatus.frx":0012
      TabIndex        =   0
      Top             =   1320
      Width           =   2055
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   7815
      Left            =   15
      Top             =   0
      Width           =   11280
   End
   Begin VB.Line Line4 
      X1              =   7680
      X2              =   7680
      Y1              =   1320
      Y2              =   960
   End
   Begin VB.Line Line3 
      X1              =   5640
      X2              =   5640
      Y1              =   1320
      Y2              =   960
   End
   Begin VB.Line Line2 
      X1              =   3600
      X2              =   3600
      Y1              =   1320
      Y2              =   960
   End
   Begin VB.Line Line1 
      X1              =   1920
      X2              =   1920
      Y1              =   1320
      Y2              =   960
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Customer No. Reg. Date            Arrival                Status                 Customer Name"
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
      Left            =   360
      TabIndex        =   6
      Top             =   960
      Width           =   10455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Registered Room Status"
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
      Left            =   360
      TabIndex        =   5
      Top             =   240
      Width           =   10815
   End
End
Attribute VB_Name = "FrmRegistrationStatus"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Form_Load()
Me.Left = MDIForm1.Width / 2 - Me.Width / 2 - 300
Me.Top = MDIForm1.Height / 2 - Me.Height / 2 - 500

On Error Resume Next
List1.Clear
List2.Clear
List3.Clear
List4.Clear
List5.Clear
SQL = "select * from CUSTOMER where ID='" & "NO" & "'"
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(CustomerPath & "\CUSTOMER.MDB")
Set MyRs = MyDb.OpenRecordset(SQL, dbOpenDynaset)

If MyRs.RecordCount <= 0 Then
Exit Sub
Else
   Do While Not MyRs.EOF
      List1.AddItem MyRs!SL
      List2.AddItem MyRs!REGDATE
      List3.AddItem MyRs!ARRIVAL
      If MyRs!ARRIVAL < Format(Now, "DD/MM/YY") Then
      List4.AddItem "EXPIRED"
      Else
      List4.AddItem "ALLIVE"
      End If
      List5.AddItem MyRs!Name
MyRs.MoveNext
    Loop
End If
MyDb.Close
End Sub

Private Sub List1_Click()
List2.ListIndex = List1.ListIndex
List3.ListIndex = List1.ListIndex
List4.ListIndex = List1.ListIndex
List5.ListIndex = List1.ListIndex

End Sub

Private Sub USStyle1_Click()
Unload Me

End Sub
