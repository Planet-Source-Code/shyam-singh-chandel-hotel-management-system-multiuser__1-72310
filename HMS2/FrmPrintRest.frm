VERSION 5.00
Object = "{97824D60-58A5-4D7C-A442-3412CD2787CD}#2.0#0"; "USSTYLE.OCX"
Begin VB.Form FrmPrintRest 
   BackColor       =   &H00008080&
   BorderStyle     =   0  'None
   Caption         =   "Print the Bill"
   ClientHeight    =   7110
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12255
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   7110
   ScaleWidth      =   12255
   ShowInTaskbar   =   0   'False
   Begin Project1.USStyle USStyle3 
      Height          =   375
      Left            =   6240
      TabIndex        =   31
      Top             =   360
      Width           =   1335
      _ExtentX        =   2355
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
   Begin Project1.USStyle USStyle2 
      Height          =   375
      Left            =   4920
      TabIndex        =   30
      Top             =   360
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Print"
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
      Left            =   720
      TabIndex        =   29
      Top             =   360
      Width           =   1575
      _ExtentX        =   2778
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
      Caption         =   "Find"
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
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   1440
      TabIndex        =   18
      Top             =   1440
      Width           =   1095
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3960
      TabIndex        =   17
      Top             =   1440
      Width           =   3615
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   615
      Left            =   1440
      MultiLine       =   -1  'True
      TabIndex        =   16
      Top             =   1920
      Width           =   6135
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   345
      Left            =   1440
      TabIndex        =   15
      Top             =   960
      Width           =   1575
   End
   Begin VB.TextBox Text5 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
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
      Left            =   3120
      TabIndex        =   14
      Top             =   960
      Width           =   4455
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   2370
      ItemData        =   "FrmPrintRest.frx":0000
      Left            =   360
      List            =   "FrmPrintRest.frx":0002
      TabIndex        =   13
      Top             =   3000
      Width           =   495
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   2370
      ItemData        =   "FrmPrintRest.frx":0004
      Left            =   840
      List            =   "FrmPrintRest.frx":0006
      TabIndex        =   12
      Top             =   3000
      Width           =   3855
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      Height          =   2370
      ItemData        =   "FrmPrintRest.frx":0008
      Left            =   4680
      List            =   "FrmPrintRest.frx":000A
      TabIndex        =   11
      Top             =   3000
      Width           =   1335
   End
   Begin VB.ListBox List4 
      Appearance      =   0  'Flat
      Height          =   2370
      ItemData        =   "FrmPrintRest.frx":000C
      Left            =   6000
      List            =   "FrmPrintRest.frx":000E
      TabIndex        =   10
      Top             =   3000
      Width           =   1575
   End
   Begin VB.TextBox Text6 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
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
      Left            =   4680
      TabIndex        =   9
      Text            =   "Total:"
      Top             =   5355
      Width           =   1335
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6000
      TabIndex        =   8
      Top             =   5355
      Width           =   1575
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   360
      TabIndex        =   7
      Top             =   5355
      Width           =   4335
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Print"
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
      Left            =   3480
      TabIndex        =   6
      Top             =   8760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
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
      Left            =   4560
      TabIndex        =   5
      Top             =   8760
      Visible         =   0   'False
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Find Bill No."
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
      Left            =   2040
      TabIndex        =   4
      Top             =   8760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.ListBox List5 
      Appearance      =   0  'Flat
      Height          =   4710
      ItemData        =   "FrmPrintRest.frx":0010
      Left            =   8400
      List            =   "FrmPrintRest.frx":0012
      Sorted          =   -1  'True
      TabIndex        =   3
      Top             =   1080
      Width           =   1695
   End
   Begin VB.ListBox List6 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   4710
      ItemData        =   "FrmPrintRest.frx":0014
      Left            =   10080
      List            =   "FrmPrintRest.frx":0016
      Sorted          =   -1  'True
      TabIndex        =   2
      Top             =   1080
      Width           =   1695
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00008080&
      Caption         =   "Original Bills"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   8280
      TabIndex        =   1
      Top             =   240
      Value           =   -1  'True
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00008080&
      Caption         =   "Duplicate Bills"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   10080
      TabIndex        =   0
      Top             =   240
      Width           =   2055
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   360
      Top             =   6480
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   7080
      Left            =   0
      Top             =   0
      Width           =   12240
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer No."
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   28
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Cash/Memo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   1
      Left            =   2880
      TabIndex        =   27
      Top             =   360
      Width           =   1815
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer Name"
      Height          =   375
      Index           =   2
      Left            =   2640
      TabIndex        =   26
      Top             =   1440
      Width           =   1455
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Address"
      Height          =   375
      Index           =   3
      Left            =   360
      TabIndex        =   25
      Top             =   1920
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No."
      Height          =   375
      Index           =   4
      Left            =   360
      TabIndex        =   24
      Top             =   960
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Date."
      Height          =   255
      Index           =   5
      Left            =   3600
      TabIndex        =   23
      Top             =   960
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   " Sl.                  Items                             Rste          Amount"
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
      TabIndex        =   22
      Top             =   2640
      Width           =   7215
   End
   Begin VB.Line Line1 
      X1              =   840
      X2              =   840
      Y1              =   3000
      Y2              =   2640
   End
   Begin VB.Line Line2 
      X1              =   4680
      X2              =   4680
      Y1              =   3000
      Y2              =   2640
   End
   Begin VB.Line Line3 
      X1              =   6000
      X2              =   6000
      Y1              =   3000
      Y2              =   2640
   End
   Begin VB.Line Line4 
      X1              =   5280
      X2              =   7320
      Y1              =   6600
      Y2              =   6600
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Signature"
      Height          =   375
      Index           =   6
      Left            =   5760
      TabIndex        =   21
      Top             =   6000
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Computer generated Billing System"
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
      Index           =   7
      Left            =   360
      TabIndex        =   20
      Top             =   6000
      Width           =   3855
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bills For Print          Printed Bills"
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
      Index           =   8
      Left            =   8400
      TabIndex        =   19
      Top             =   840
      Width           =   3015
   End
End
Attribute VB_Name = "FrmPrintRest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyDb As Database, MyRs As Recordset

Private Sub Command1_Click()
On Error Resume Next
Me.BackColor = vbWhite
Shape1.Visible = False
Image1.Visible = False
SQL = "select * from RESTO where BILLNO='" & Text4.Text & "'"
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(RestoPath & "\RESTO.MDB")
Set MyRs = MyDb.OpenRecordset(SQL, dbOpenDynaset)
Do While Not MyRs.EOF
   MyRs.Edit
     MyRs!PRINT = "DONE"
   MyRs.Update
   MyRs.MoveNext
Loop
MyDb.Close
USStyle1.Visible = False
USStyle3.Visible = False
'Command3.Visible = False
Me.Width = 8055
USStyle2.Visible = False
FrmPrintRest.PrintForm
USStyle2.Visible = True
USStyle2.Enabled = False
USStyle1.Visible = True
USStyle3.Visible = True
Me.Width = 12255
Me.BackColor = &H8080&
Shape1.Visible = True
Image1.Visible = True
Form_Load
'Command1.Default = False
'Command1.Enabled = False
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Command3_Click()
BILLENTRY
End Sub


Private Sub Form_Load()
On Error Resume Next
Me.Left = MDIForm1.Width / 2 - Me.Width / 2 - 300
Me.Top = MDIForm1.Height / 2 - Me.Height / 2 - 500

List5.Clear
 List6.Clear
 
SQL = "select * from RESTO where PRINT='" & "NOTDONE" & "'"
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(RestoPath & "\RESTO.MDB")
Set MyRs = MyDb.OpenRecordset(SQL, dbOpenDynaset)
   Do While Not MyRs.EOF
     List5.AddItem MyRs!BILLNO
    MyRs.MoveNext
    Loop
 SQL2 = "select * from RESTO where PRINT='" & "DONE" & "'"
Set MyRs = MyDb.OpenRecordset(SQL2, dbOpenDynaset)
   Do While Not MyRs.EOF
      List6.AddItem MyRs!BILLNO
   MyRs.MoveNext
    Loop
   MyDb.Close
End Sub

Private Sub BILLENTRY()
List1.Clear
List2.Clear
List3.Clear
List4.Clear
Text1 = ""
Text2 = ""
Text3 = ""
Text7 = ""

Dim Amt, SL
SL = 0
On Error Resume Next
SQL = "select * from RESTO where BILLNO='" & Text4.Text & "'"
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(RestoPath & "\RESTO.MDB")
Set MyRs = MyDb.OpenRecordset(SQL, dbOpenDynaset)
   Do While Not MyRs.EOF
      Text1.Text = MyRs!CUSTOMERNO
      Text2.Text = MyRs!CUSTOMERNAME
      Text3.Text = MyRs!ADDRESS
      SL = SL + 1
      List1.AddItem SL
      List2.AddItem MyRs!ITEMNAME
      List3.AddItem MyRs!Rate & "/="
      Amt = MyRs!AM0UNT
      List4.AddItem Amt & "/="
      Text7.Text = Val(Text7.Text) + Amt & "/="
    MyRs.MoveNext
    Loop
    MyDb.Close
End Sub

Private Sub List5_Click()
Command1.Enabled = True
USStyle2.Enabled = True
Text4.Text = List5.Text
BILLENTRY
Command1.Default = True
End Sub

Private Sub List6_Click()
Command1.Enabled = True
USStyle2.Enabled = True
Text4.Text = List6.Text
BILLENTRY
Command1.Default = True
End Sub

Private Sub Option1_Click()
If Option1.Value = True Then
   List5.Enabled = True
   List6.Enabled = False
   List1.Clear
    List2.Clear
    List3.Clear
    List4.Clear
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text7 = ""
   End If
   
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
   List6.Enabled = True
   List5.Enabled = False
   End If
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
BILLENTRY
Command1.Default = True
End If

End Sub

Private Sub Timer1_Timer()
Text5.Text = "Date: " & Format(Now, "DD-MMMM=YYYY") & "    Time: " & Format(Now, "HH:MM AM/PM")
End Sub

Private Sub USStyle1_Click()
BILLENTRY
End Sub

Private Sub USStyle2_Click()
Command1_Click
End Sub

Private Sub USStyle3_Click()
Unload Me

End Sub
