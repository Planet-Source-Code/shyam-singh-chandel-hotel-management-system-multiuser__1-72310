VERSION 5.00
Object = "{97824D60-58A5-4D7C-A442-3412CD2787CD}#2.0#0"; "USSTYLE.OCX"
Begin VB.Form FrmPrintBill 
   BackColor       =   &H00008080&
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   13365
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7365
   ScaleWidth      =   13365
   ShowInTaskbar   =   0   'False
   Begin Project1.USStyle USStyle2 
      Height          =   375
      Left            =   11400
      TabIndex        =   36
      Top             =   240
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   11.25
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
   Begin Project1.USStyle USStyle1 
      Height          =   375
      Left            =   480
      TabIndex        =   34
      Top             =   120
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   661
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Print"
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
   Begin VB.CommandButton Command2 
      Caption         =   "Print"
      Enabled         =   0   'False
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
      Left            =   13200
      TabIndex        =   33
      Top             =   6000
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1920
      TabIndex        =   23
      Top             =   1080
      Width           =   5295
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   855
      Left            =   1920
      MultiLine       =   -1  'True
      TabIndex        =   22
      Top             =   1560
      Width           =   5295
   End
   Begin VB.TextBox Text3 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   9000
      TabIndex        =   21
      Top             =   2040
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5160
      TabIndex        =   20
      Top             =   600
      Width           =   5535
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2955
      ItemData        =   "FrmPrint.frx":0000
      Left            =   480
      List            =   "FrmPrint.frx":0002
      TabIndex        =   19
      Top             =   2880
      Width           =   495
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   2955
      ItemData        =   "FrmPrint.frx":0004
      Left            =   960
      List            =   "FrmPrint.frx":0006
      TabIndex        =   18
      Top             =   2880
      Width           =   7095
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      Height          =   2955
      ItemData        =   "FrmPrint.frx":0008
      Left            =   8040
      List            =   "FrmPrint.frx":000A
      TabIndex        =   17
      Top             =   2880
      Width           =   1095
   End
   Begin VB.ListBox List4 
      Appearance      =   0  'Flat
      Height          =   2955
      ItemData        =   "FrmPrint.frx":000C
      Left            =   9120
      List            =   "FrmPrint.frx":000E
      TabIndex        =   16
      Top             =   2880
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   8040
      TabIndex        =   15
      Text            =   "Total: "
      Top             =   5805
      Width           =   1095
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   405
      Left            =   9120
      TabIndex        =   14
      Top             =   5805
      Width           =   1575
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   480
      TabIndex        =   13
      Top             =   5805
      Width           =   7580
   End
   Begin VB.TextBox Text8 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   12
      Top             =   1080
      Width           =   1695
   End
   Begin VB.TextBox Text9 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   9000
      TabIndex        =   11
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1920
      TabIndex        =   10
      Top             =   600
      Width           =   1935
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   6120
      Top             =   3600
   End
   Begin VB.TextBox Text13 
      Height          =   405
      Left            =   3840
      TabIndex        =   8
      Top             =   11160
      Width           =   2055
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   3840
      TabIndex        =   7
      Top             =   10680
      Width           =   2055
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   3840
      TabIndex        =   6
      Top             =   10200
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
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
      Left            =   13200
      TabIndex        =   5
      Top             =   120
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox List9 
      Appearance      =   0  'Flat
      Height          =   2175
      ItemData        =   "FrmPrint.frx":0010
      Left            =   360
      List            =   "FrmPrint.frx":0012
      TabIndex        =   4
      Top             =   10515
      Width           =   495
   End
   Begin VB.ListBox List8 
      Appearance      =   0  'Flat
      Height          =   2175
      ItemData        =   "FrmPrint.frx":0014
      Left            =   840
      List            =   "FrmPrint.frx":0016
      TabIndex        =   3
      Top             =   10515
      Width           =   7095
   End
   Begin VB.ListBox List7 
      Appearance      =   0  'Flat
      Height          =   2175
      ItemData        =   "FrmPrint.frx":0018
      Left            =   7920
      List            =   "FrmPrint.frx":001A
      TabIndex        =   2
      Top             =   10515
      Width           =   1095
   End
   Begin VB.ListBox List6 
      Appearance      =   0  'Flat
      Height          =   2175
      ItemData        =   "FrmPrint.frx":001C
      Left            =   9000
      List            =   "FrmPrint.frx":001E
      TabIndex        =   1
      Top             =   10515
      Width           =   1575
   End
   Begin VB.ListBox List5 
      Appearance      =   0  'Flat
      Height          =   5100
      ItemData        =   "FrmPrint.frx":0020
      Left            =   11400
      List            =   "FrmPrint.frx":0022
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   960
      Width           =   1575
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer No."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   11400
      TabIndex        =   35
      Top             =   720
      Width           =   1575
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Height          =   7335
      Left            =   0
      Top             =   0
      Width           =   13335
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Customer Name"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   480
      TabIndex        =   32
      Top             =   1080
      Width           =   1455
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Address"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   855
      Left            =   480
      TabIndex        =   31
      Top             =   1560
      Width           =   1455
   End
   Begin VB.Label Label3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Total Days:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7440
      TabIndex        =   30
      Top             =   2040
      Width           =   1575
   End
   Begin VB.Label Label4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "This is Computer generated bill.                                                            Signature"
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
      Height          =   495
      Left            =   480
      TabIndex        =   29
      Top             =   6480
      Width           =   10215
   End
   Begin VB.Label Label5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sl.                                        Items                                                   Rate        Amount"
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
      Left            =   480
      TabIndex        =   28
      Top             =   2520
      Width           =   10215
   End
   Begin VB.Line Line1 
      X1              =   960
      X2              =   960
      Y1              =   2880
      Y2              =   2520
   End
   Begin VB.Line Line2 
      X1              =   8040
      X2              =   8040
      Y1              =   2880
      Y2              =   2520
   End
   Begin VB.Line Line3 
      X1              =   9120
      X2              =   9120
      Y1              =   2880
      Y2              =   2520
   End
   Begin VB.Label Label6 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Cash/Memo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4080
      TabIndex        =   27
      Top             =   120
      Width           =   1935
   End
   Begin VB.Label Label7 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Check In Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7440
      TabIndex        =   26
      Top             =   1080
      Width           =   1575
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Check Out Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   7440
      TabIndex        =   25
      Top             =   1560
      Width           =   1575
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Customer No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   480
      TabIndex        =   24
      Top             =   600
      Width           =   1455
   End
   Begin VB.Line Line4 
      X1              =   8160
      X2              =   10560
      Y1              =   7080
      Y2              =   7080
   End
   Begin VB.Label Label10 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Customer No"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   11400
      TabIndex        =   9
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "FrmPrintBill"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyDb As Database, MyRs As Recordset
Dim Amt

Private Sub BILLENTRY()
List1.Clear
List2.Clear
List3.Clear
List4.Clear
Text1 = ""
Text2 = ""
Text3 = ""
Text6 = ""
Text7 = ""

Dim Amt, SL, RC, Rs

SL = 0
On Error Resume Next
SQL = "select * from PRINTDB where ID='" & Text10.Text & "'"
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(PrintPath & "\PRINT.MDB")
Set MyRs = MyDb.OpenRecordset(SQL, dbOpenDynaset)
      Text1.Text = MyRs!Name
      Text2.Text = MyRs!ADDRESS
      Text3.Text = MyRs!NOOFDAYS
      Text9.Text = MyRs!checkoutdate
      Text8.Text = MyRs!checkindate
      
   SL = SL + 1
        List1.AddItem SL
        List2.AddItem "LODGING"
        Amt = MyRs!lodging
        List3.AddItem MyRs!roomcharges & "/="
        List4.AddItem Amt & "/="
   Text11.Text = Amt
   
   Amt = ""
   SL = SL + 1
        List1.AddItem SL
        List2.AddItem "ROOM SERVICE"
        Amt = Val(MyRs!NOOFDAYS) * 50
        List3.AddItem "50" & "/="
        List4.AddItem Amt & "/="
   Text12.Text = Amt
   
    SL = SL + 1
        List1.AddItem SL
        List2.AddItem "FOODING CHARGES"
        List3.AddItem MyRs!FOODING & "/="
        Amt = MyRs!FOODING
        List4.AddItem Amt & "/="
    Text13.Text = Amt
   
    Text6.Text = Val(Text11.Text) + Val(Text12.Text) + Val(Text13.Text) & "/="
    
End Sub

Private Sub Command1_Click()
Unload Me

End Sub

Private Sub Command2_Click()
On Error Resume Next
Me.BackColor = vbWhite
Me.Width = 10965
Image1.Visible = False
Shape1.Visible = False
'Command2.Visible = False
USStyle1.Visible = False
Me.PrintForm
'Command2.Visible = True
USStyle1.Visible = True
USStyle1.Enabled = False
Image1.Visible = True
Shape1.Visible = True
Me.Width = 13365
Me.BackColor = &H8080&
Do While Not MyRs.EOF
    MyRs.Edit
        MyRs!printstatus = "DONE"
    MyRs.Update
    MyRs.MoveNext
    Loop
    List5.Clear
    Form_Load
    List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text6 = ""
    Text7 = ""
    Command2.Enabled = False
End Sub

Private Sub Form_Load()
Me.Left = MDIForm1.Width / 2 - Me.Width / 2 - 300
Me.Top = MDIForm1.Height / 2 - Me.Height / 2 - 500

SQL = "select * from PRINTDB where PRINTSTATUS='" & "NOTDONE" & "'"
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(PrintPath & "\PRINT.MDB")
Set MyRs = MyDb.OpenRecordset(SQL, dbOpenDynaset)
   Do While Not MyRs.EOF
         List5.AddItem MyRs!ID
    MyRs.MoveNext
    Loop
End Sub

Private Sub List5_Click()
Command2.Enabled = True
USStyle1.Enabled = True
Text10.Text = List5.Text
BILLENTRY
End Sub

Private Sub Timer1_Timer()
Text4.Text = "Date:- " & Format(Now, "DD-MMMM-YY") & "    Time:- " & Format(Now, "HH:MM AM/PM")
End Sub

Private Sub USStyle1_Click()
Command2_Click
End Sub

Private Sub USStyle2_Click()
Unload Me

End Sub
