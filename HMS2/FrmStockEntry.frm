VERSION 5.00
Object = "{97824D60-58A5-4D7C-A442-3412CD2787CD}#2.0#0"; "USSTYLE.OCX"
Begin VB.Form FrmStockEntry 
   BackColor       =   &H00008080&
   BorderStyle     =   0  'None
   Caption         =   "Stock Entry"
   ClientHeight    =   6990
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7005
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   6990
   ScaleWidth      =   7005
   ShowInTaskbar   =   0   'False
   Begin Project1.USStyle USStyle4 
      Height          =   495
      Left            =   2880
      TabIndex        =   20
      Top             =   5880
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
      Style           =   2
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
      Left            =   2880
      TabIndex        =   19
      Top             =   5280
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
      Caption         =   "Delete Stock"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   4210752
      Style           =   2
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
      Left            =   2880
      TabIndex        =   18
      Top             =   4680
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
      Caption         =   "Edit Stock Record"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   4210752
      Style           =   2
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
      Left            =   2880
      TabIndex        =   17
      Top             =   4080
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
      Caption         =   "Update Stock"
      IconHighLiteColor=   0
      CaptionHighLiteColor=   0
      ForeColor       =   4210752
      Style           =   2
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
      BackColor       =   &H00808080&
      Caption         =   "Delete Stock"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   4920
      Width           =   3255
   End
   Begin VB.TextBox Text5 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   3840
      TabIndex        =   4
      Top             =   3480
      Width           =   2295
   End
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   5520
      Width           =   3255
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Update Stock"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   3720
      Width           =   3255
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      BackColor       =   &H0000C000&
      Height          =   4905
      Left            =   600
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   1560
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3840
      TabIndex        =   1
      Top             =   2040
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   3840
      TabIndex        =   0
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   3840
      TabIndex        =   2
      Top             =   2520
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Appearance      =   0  'Flat
      Height          =   405
      Left            =   3840
      TabIndex        =   3
      Top             =   3000
      Width           =   2295
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H00808080&
      Caption         =   "Edit Stock Record"
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
      Left            =   7080
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4320
      Width           =   3255
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000C0C0&
      BorderWidth     =   3
      Height          =   6975
      Left            =   20
      Top             =   20
      Width           =   6975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Menu Stock Entry"
      BeginProperty Font 
         Name            =   "Times New Roman"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   735
      Index           =   3
      Left            =   120
      TabIndex        =   15
      Top             =   480
      Width           =   6735
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Stock Entry By"
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
      Left            =   2520
      TabIndex        =   14
      Top             =   3480
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item No. List"
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
      Left            =   600
      TabIndex        =   13
      Top             =   1320
      Width           =   1455
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Item Name"
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
      Left            =   2520
      TabIndex        =   12
      Top             =   2040
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Item No."
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
      Left            =   2520
      TabIndex        =   11
      Top             =   1560
      Width           =   1215
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "Opening Stock"
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
      Left            =   2520
      TabIndex        =   10
      Top             =   3000
      Width           =   1335
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Price"
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
      Left            =   2520
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
End
Attribute VB_Name = "FrmStockEntry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyDb As Database, MyRs As Recordset

Private Sub Command1_Click()
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(ItemsPath & "\ITEMS.MDB")
Set MyRs = MyDb.OpenRecordset("ITEMS", dbOpenDynaset)
MyRs.AddNew
    MyRs!ITEMNO = Text1.Text
    MyRs!ITEMNAME = Text2.Text
    MyRs!Rate = Text3.Text
    MyRs!openingstock = Text4.Text
    MyRs!closingstock = Text4.Text
    MyRs!STOCKENTRYPERSON = Text5.Text
 MyRs.Update

 MsgBox "RECORD OF STOCK IS UPDATED"
 Text1 = ""
 Text2 = ""
 Text3 = ""
 Text4 = ""
 Text5 = ""
 Form_Load
Command3.Enabled = False
Command4.Enabled = False
Command1.Enabled = False
End Sub

Private Sub Command2_Click()
 
 Unload Me
 
End Sub

Private Sub Command3_Click()

MyRs.Edit
    MyRs!ITEMNO = Text1.Text
    MyRs!ITEMNAME = Text2.Text
    MyRs!Rate = Text3.Text
    MyRs!openingstock = Text4.Text
    MyRs!STOCKENTRYPERSON = Text5.Text
 MyRs.Update
 MyRs.Close
 MyDb.Close
 MsgBox "RECORD OF STOCK IS UPDATED"
 Text1 = ""
 Text2 = ""
 Text3 = ""
 Text4 = ""
 Text5 = ""
 Command3.Enabled = False
Command4.Enabled = False
Command1.Enabled = False
 Form_Load

End Sub

Private Sub Command4_Click()
res = MsgBox("Are you sure that you want to delete the record of '" & Text2.Text & "'.", vbYesNo + vbQuestion)
If res = vbYes Then
MyRs.Delete
Form_Load
Else
Exit Sub
End If
 Text1 = ""
 Text2 = ""
 Text3 = ""
 Text4 = ""
 Text5 = ""
Command3.Enabled = False
Command4.Enabled = False
Command1.Enabled = False
End Sub

Private Sub Form_Load()
On Error Resume Next
Me.Left = MDIForm1.Width / 2 - Me.Width / 2 - 300 '4000
Me.Top = MDIForm1.Height / 2 - Me.Height / 2 - 500

List1.Clear
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(ItemsPath & "\ITEMS.MDB")
Set MyRs = MyDb.OpenRecordset("ITEMS", dbOpenDynaset)
If MyRs.RecordCount <= 0 Then
Exit Sub
Else
   Do While Not MyRs.EOF
      List1.AddItem MyRs!ITEMNO
    MyRs.MoveNext
    Loop
End If
MyDb.Close
 Text1.SetFocus
End Sub

Private Sub List1_Click()
Command3.Enabled = True
Command4.Enabled = True
USStyle2.Enabled = True
USStyle3.Enabled = True

SQL = "select * from ITEMS where ITEMNO='" & List1.Text & "'"
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(ItemsPath & "\ITEMS.MDB")
Set MyRs = MyDb.OpenRecordset(SQL, dbOpenDynaset)

   Text1.Text = MyRs!ITEMNO
   Text2.Text = MyRs!ITEMNAME
   Text3.Text = MyRs!Rate
   Text4.Text = MyRs!openingstock
   Text5.Text = MyRs!STOCKENTRYPERSON

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
End If
End Sub

Private Sub Text5_Change()
Command1.Enabled = True
End Sub

Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  USStyle1.Enabled = True
  Command1.Enabled = True
  Text5.SetFocus
  
End If
End Sub
Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  Command1.Enabled = True
  'Command1_Click
  USStyle1_Click
  Command1.Enabled = False
  Text1.SetFocus

End If
End Sub

Private Sub USStyle1_Click()
Command1_Click
End Sub

Private Sub USStyle2_Click()
Command3_Click
End Sub

Private Sub USStyle3_Click()
Command4_Click
End Sub

Private Sub USStyle4_Click()
Unload Me

End Sub
