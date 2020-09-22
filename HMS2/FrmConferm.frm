VERSION 5.00
Object = "{97824D60-58A5-4D7C-A442-3412CD2787CD}#2.0#0"; "USSTYLE.OCX"
Begin VB.Form FrmConferm 
   BackColor       =   &H00008080&
   BorderStyle     =   0  'None
   Caption         =   "Registration Confermation"
   ClientHeight    =   7380
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11820
   Icon            =   "FrmConferm.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7380
   ScaleWidth      =   11820
   ShowInTaskbar   =   0   'False
   Begin Project1.USStyle USStyle2 
      Height          =   495
      Left            =   8880
      TabIndex        =   44
      Top             =   240
      Width           =   2415
      _ExtentX        =   4260
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
      Style           =   7
      Checked         =   0   'False
      ColorButtonHover=   41120
      ColorButtonUp   =   32896
      ColorButtonDown =   49344
      BorderBrightness=   0
      ColorBright     =   65535
      DisplayHand     =   0   'False
      ColorScheme     =   6
   End
   Begin Project1.USStyle USStyle1 
      Height          =   855
      Left            =   480
      TabIndex        =   43
      Top             =   6000
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1508
      Enabled         =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Times New Roman"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Confirm Registration"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   8454143
      Style           =   3
      Checked         =   0   'False
      ColorButtonHover=   40960
      ColorButtonUp   =   32768
      ColorButtonDown =   49152
      BorderBrightness=   0
      ColorBright     =   65280
      DisplayHand     =   0   'False
      ColorScheme     =   5
   End
   Begin VB.CommandButton Command7 
      BackColor       =   &H00FFFF80&
      Caption         =   "Change Arrival / Checkout Date"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   3360
      Style           =   1  'Graphical
      TabIndex        =   42
      Top             =   9840
      Width           =   2535
   End
   Begin VB.CommandButton Command6 
      BackColor       =   &H00FFFF80&
      Caption         =   "Change Type of Room"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   6120
      Style           =   1  'Graphical
      TabIndex        =   41
      Top             =   9840
      Width           =   2415
   End
   Begin VB.TextBox Text16 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   2100
      Left            =   1560
      TabIndex        =   40
      Top             =   3300
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.ListBox List1 
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
      Height          =   2130
      ItemData        =   "FrmConferm.frx":0CCA
      Left            =   1680
      List            =   "FrmConferm.frx":0CCC
      TabIndex        =   39
      Top             =   3285
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.TextBox Text15 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      TabIndex        =   37
      Top             =   6480
      Width           =   1935
   End
   Begin VB.TextBox Text14 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   6600
      TabIndex        =   35
      Top             =   5880
      Width           =   1935
   End
   Begin VB.ListBox List3 
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
      Height          =   2130
      ItemData        =   "FrmConferm.frx":0CCE
      Left            =   360
      List            =   "FrmConferm.frx":0CD0
      TabIndex        =   34
      Top             =   3285
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton Command5 
      BackColor       =   &H0000C0C0&
      Caption         =   "Confirm Registration"
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
      Height          =   855
      Left            =   480
      Style           =   1  'Graphical
      TabIndex        =   33
      Top             =   9840
      Width           =   2655
   End
   Begin VB.TextBox Text13 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   6600
      TabIndex        =   30
      Top             =   5400
      Width           =   1935
   End
   Begin VB.TextBox Text12 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   26
      Top             =   5400
      Width           =   1815
   End
   Begin VB.TextBox Text11 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2520
      TabIndex        =   25
      Top             =   5400
      Width           =   2055
   End
   Begin VB.TextBox Text10 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   360
      TabIndex        =   24
      Top             =   5400
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   13
      Top             =   1200
      Width           =   1935
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   2400
      TabIndex        =   12
      Top             =   1200
      Width           =   6135
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1095
      Left            =   360
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   11
      Top             =   2040
      Width           =   8175
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   10
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   2520
      TabIndex        =   9
      Top             =   3600
      Width           =   6015
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   360
      TabIndex        =   8
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   4560
      Width           =   2055
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   6
      Top             =   4560
      Width           =   2295
   End
   Begin VB.TextBox Text9 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   7080
      TabIndex        =   5
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton Command1 
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
      Left            =   960
      TabIndex        =   4
      Top             =   11160
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
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
      Left            =   2160
      TabIndex        =   3
      Top             =   11160
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
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
      Left            =   3360
      TabIndex        =   2
      Top             =   11160
      Width           =   975
   End
   Begin VB.CommandButton Command4 
      BackColor       =   &H000000FF&
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
      Height          =   855
      Left            =   8640
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9840
      Width           =   2655
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   5685
      Left            =   8880
      Sorted          =   -1  'True
      TabIndex        =   0
      Top             =   1200
      Width           =   2415
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Height          =   7335
      Left            =   15
      Top             =   15
      Width           =   11760
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Balance"
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
      Left            =   5280
      TabIndex        =   38
      Top             =   6480
      Width           =   1335
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Advance"
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
      Left            =   5280
      TabIndex        =   36
      Top             =   5880
      Width           =   1215
   End
   Begin VB.Label Label15 
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
      TabIndex        =   32
      Top             =   840
      Width           =   1455
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Total Rent"
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
      Left            =   6600
      TabIndex        =   31
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "No. of Days"
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
      Left            =   7080
      TabIndex        =   29
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Rent per Day"
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
      Left            =   4680
      TabIndex        =   28
      Top             =   5040
      Width           =   1815
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Out Date"
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
      TabIndex        =   27
      Top             =   5040
      Width           =   2175
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
      Left            =   2400
      TabIndex        =   23
      Top             =   840
      Width           =   975
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
      Left            =   360
      TabIndex        =   22
      Top             =   1680
      Width           =   1575
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
      Left            =   360
      TabIndex        =   21
      Top             =   3240
      Width           =   1335
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
      Left            =   2520
      TabIndex        =   20
      Top             =   3240
      Width           =   975
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
      Left            =   360
      TabIndex        =   19
      Top             =   840
      Width           =   1455
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
      Left            =   360
      TabIndex        =   18
      Top             =   4200
      Width           =   2055
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
      Left            =   2520
      TabIndex        =   17
      Top             =   4200
      Width           =   2055
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Room No."
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
      Left            =   360
      TabIndex        =   16
      Top             =   5040
      Width           =   1335
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
      Left            =   4680
      TabIndex        =   15
      Top             =   4200
      Width           =   2175
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Registration Confirmation"
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
      TabIndex        =   14
      Top             =   120
      Width           =   5655
   End
End
Attribute VB_Name = "FrmConferm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyDb As Database, MyRs As Recordset
Dim Db As Database, Rs As Recordset
Private Sub UpdateDate()
    Dim Date1 As Date, Date2 As Date
    Dim Date3 As Date, Date4 As Date
    Dim Years As Integer, Months As Integer, Days As Integer
    'set dates
    Date3 = Now
    'set the difference
   ' Date4 = DateAdd("yyyy", 0, DateAdd("m", 0, DateAdd("d", Val(Text1.Text), Date3)))
    Date4 = DateAdd("d", Val(Text9.Text), Date3)
   
    'view the dates for user
    Text11.Text = Format(Date4, "dd/mm/yy")
    'we can't have negative difference: thus check which date is bigger and which is smaller
    If Date3 < Date4 Then
        Date1 = Date3
        Date2 = Date4
    Else
        Date1 = Date4
        Date2 = Date3
    End If
    'view DateDiff results to user
    DateDiffLabel = DateDiff("yyyy", Date1, Date2) & " year(s), " & DateDiff("m", Date1, Date2) & " month(s) and " & DateDiff("d", Date1, Date2) & " day(s)."
    'get years
    Years = DateDiff("yyyy", Date1, Date2)
    'get months and decrease by one if there can't be one complete month by according the days
    Months = DateDiff("m", Date1, Date2) + (Day(Date1) > Day(Date2))
    'decrease years if necessary for the same reason the months were decreased
    Years = Years + ((Months - Years * 12) < 0)
    'rip out extra months
    Months = Months Mod 12
    'get day difference
    Days = DateDiff("d", Date1, Date2) - DateDiff("d", Date1, DateAdd("yyyy", Years, DateAdd("m", Months, Date1)))
    'view the difference for user
    DateDifference = Years & " year(s), " & Months & " month(s) and " & Days & " day(s)."
End Sub

Private Sub UpdateDate2()
    Dim Date1 As Date, Date2 As Date
    Dim Date3 As Date, Date4 As Date
    Dim Years As Integer, Months As Integer, Days As Integer
    'set dates
    Date3 = Now
    'set the difference
   ' Date4 = DateAdd("yyyy", 0, DateAdd("m", 0, DateAdd("d", Val(Text1.Text), Date3)))
    Date4 = DateAdd("d", Val(Text9.Text), Date3)
   
    'view the dates for user
    Text7.Text = Format(Date3, "dd/mm/yy")
    Text11.Text = Format(Date4, "dd/mm/yy")
    'we can't have negative difference: thus check which date is bigger and which is smaller
    If Date3 < Date4 Then
        Date1 = Date3
        Date2 = Date4
    Else
        Date1 = Date4
        Date2 = Date3
    End If
    'view DateDiff results to user
    DateDiffLabel = DateDiff("yyyy", Date1, Date2) & " year(s), " & DateDiff("m", Date1, Date2) & " month(s) and " & DateDiff("d", Date1, Date2) & " day(s)."
    'get years
    Years = DateDiff("yyyy", Date1, Date2)
    'get months and decrease by one if there can't be one complete month by according the days
    Months = DateDiff("m", Date1, Date2) + (Day(Date1) > Day(Date2))
    'decrease years if necessary for the same reason the months were decreased
    Years = Years + ((Months - Years * 12) < 0)
    'rip out extra months
    Months = Months Mod 12
    'get day difference
    Days = DateDiff("d", Date1, Date2) - DateDiff("d", Date1, DateAdd("yyyy", Years, DateAdd("m", Months, Date1)))
    'view the difference for user
    DateDifference = Years & " year(s), " & Months & " month(s) and " & Days & " day(s)."
End Sub


Private Sub Command4_Click()
Unload Me
End Sub

Private Sub Command5_Click()
MyRs.Edit
     MyRs!ARRIVAL = Text7.Text
     MyRs!checkindate = Format(Date, "DD/MM/YY")
     MyRs!CHECKINTIME = Format(Now, "HH:MM:SS AM/PM")
     MyRs!checkoutdate = Text11.Text
     MyRs!CHECKOUTTIME = Format(Now, "HH:MM:SS AM/PM")
     MyRs!ROOMCHARGES = Text12.Text
     MyRs!ADVANCE = Text14.Text
     MyRs!BALANCE = Text15.Text
     MyRs!ROOMNO = Text10.Text
     MyRs!ID = "YES"
     MyRs!CHECKOUTSTATUS = "NOTDONE"
MyRs.Update
   SQL = "select * from ROOMS where ROOMNO='" & List3.Text & "'"
Set Db = DBEngine.Workspaces(0).OpenDatabase(RoomsPath & "\ROOMS.MDB")
Set Rs = Db.OpenRecordset(SQL, dbOpenDynaset)

Rs.Edit
Rs!Status = "OCCUPIED"
Rs.Update
Rs.Close
Db.Close
MsgBox "Registration is Confermed.", vbOKOnly + vbInformation
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
     Text10.Text = ""
     Text11.Text = ""
     Text12.Text = ""
     Text13.Text = ""
     Text14.Text = ""
     Text15.Text = ""
Command5.Enabled = False
Form_Load

End Sub

Private Sub Command6_Click()
FrmRegister.Show
Unload Me
End Sub

Private Sub Command7_Click()
UpdateDate2

End Sub

Private Sub Form_Load()
Dim D
'USStyle1.Style = Mac_Variation
Me.Left = MDIForm1.Width / 2 - Me.Width / 2 - 300
Me.Top = MDIForm1.Height / 2 - Me.Height / 2 - 500

On Error Resume Next
List2.Clear
D = Format(Now, "dd/mm/yy")
Text6.Text = D
Text7.Text = D

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
MyDb.Close

End Sub



Private Sub List1_Click()
List1.ListIndex = List3.ListIndex
List1.Visible = False
List3.Visible = False
Text10.Text = List3.Text
List3_Click
End Sub

Private Sub List2_Click()
On Error Resume Next
     List3.Clear
     List1.Clear
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
     Text10.Text = ""
     Text11.Text = ""
     Text12.Text = ""
     Text13.Text = ""
     Text14.Text = ""
     Text15.Text = ""
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
   Text7.Text = MyRs!checkindate
   Text10.Text = MyRs!ROOMNO
   Text11.Text = MyRs!checkoutdate
   Text12.Text = MyRs!ROOMCHARGES
   Text14.Text = MyRs!ADVANCE
   Text15.Text = MyRs!BALANCE
   
   SQL = "select * from ROOMS where TYPEOFROOM='" & Text8.Text & "'"
Set Db = DBEngine.Workspaces(0).OpenDatabase(RoomsPath & "\ROOMS.MDB")
Set Rs = Db.OpenRecordset(SQL, dbOpenDynaset)
Do While Not Rs.EOF
        List3.AddItem Rs!ROOMNO
        List1.AddItem Rs!Status
Rs.MoveNext
Loop
Dim DL
'DL = Format(Date, "dd") + Val(Text9.Text) - 1
'Text11.Text = DL & Format(Date, "/mm/yy")

UpdateDate
End Sub

Private Sub List3_Click()
On Error Resume Next
List1.ListIndex = List3.ListIndex
Text10.Text = List3.Text
List3.Visible = False
List1.Visible = False
Text16.Visible = False
 SQL = "select * from ROOMS where TYPEOFROOM='" & Text8.Text & "'"
Set Db = DBEngine.Workspaces(0).OpenDatabase(RoomsPath & "\ROOMS.MDB")
Set Rs = Db.OpenRecordset(SQL, dbOpenDynaset)
  Text12.Text = Rs!Rate
  
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text2.SetFocus
 End If
End Sub


Private Sub Text10_Click()
List3.Visible = True
List1.Visible = True
Text16.Visible = True
End Sub


Private Sub Text12_Change()
 Text13.Text = Val(Text12.Text) * Val(Text9.Text)
End Sub

Private Sub Text14_Change()
Text15.Text = Val(Text13.Text) - Val(Text14.Text)
Command5.Enabled = True
USStyle1.Enabled = True
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text3.SetFocus
 End If
End Sub
Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text4.SetFocus
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
    Text8.SetFocus
    
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
'   Dim DL
'DL = Format(Date, "dd") + Val(Text9.Text) - 1
'Text11.Text = DL & Format(Date, "/mm/yy")
'Text10.SetFocus
 End If
End Sub


Private Sub USStyle1_Click()
Command5_Click
End Sub

Private Sub USStyle2_Click()
Unload Me

End Sub
