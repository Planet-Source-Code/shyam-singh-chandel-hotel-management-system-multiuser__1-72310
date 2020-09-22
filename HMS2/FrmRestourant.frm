VERSION 5.00
Object = "{97824D60-58A5-4D7C-A442-3412CD2787CD}#2.0#0"; "USSTYLE.OCX"
Begin VB.Form FrmRestourant 
   BackColor       =   &H00008080&
   BorderStyle     =   0  'None
   Caption         =   "Restourant"
   ClientHeight    =   7245
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   11565
   Icon            =   "FrmRestourant.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7245
   ScaleWidth      =   11565
   ShowInTaskbar   =   0   'False
   Begin Project1.USStyle USStyle3 
      Height          =   495
      Left            =   4560
      TabIndex        =   73
      Top             =   6360
      Width           =   1815
      _ExtentX        =   3201
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
      Left            =   2520
      TabIndex        =   72
      Top             =   6360
      Width           =   1935
      _ExtentX        =   3413
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
      Caption         =   "OK"
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
      Height          =   495
      Left            =   360
      TabIndex        =   71
      Top             =   6360
      Width           =   2055
      _ExtentX        =   3625
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
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Find Customer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4455
      Left            =   360
      TabIndex        =   41
      Top             =   1680
      Visible         =   0   'False
      Width           =   10815
      Begin VB.CommandButton Command4 
         BackColor       =   &H00808080&
         Caption         =   "Done"
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
         Left            =   8640
         Style           =   1  'Graphical
         TabIndex        =   70
         Top             =   3960
         Width           =   1935
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4560
         TabIndex        =   55
         Top             =   3720
         Width           =   1815
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6480
         TabIndex        =   54
         Top             =   3720
         Width           =   1935
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   240
         TabIndex        =   53
         Top             =   840
         Width           =   1935
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2280
         TabIndex        =   52
         Top             =   840
         Width           =   6135
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         Height          =   1095
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   51
         Top             =   1560
         Width           =   5055
      End
      Begin VB.TextBox Text14 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5400
         TabIndex        =   50
         Top             =   1560
         Width           =   3015
      End
      Begin VB.TextBox Text15 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5400
         TabIndex        =   49
         Top             =   2280
         Width           =   3015
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   240
         TabIndex        =   48
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         TabIndex        =   47
         Top             =   3000
         Width           =   2055
      End
      Begin VB.TextBox Text18 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4560
         TabIndex        =   46
         Top             =   3000
         Width           =   2295
      End
      Begin VB.TextBox Text19 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6960
         TabIndex        =   45
         Top             =   3000
         Width           =   1455
      End
      Begin VB.TextBox Text20 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   240
         TabIndex        =   44
         Top             =   3720
         Width           =   2055
      End
      Begin VB.TextBox Text21 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2400
         TabIndex        =   43
         Top             =   3720
         Width           =   2055
      End
      Begin VB.ListBox List6 
         Appearance      =   0  'Flat
         Height          =   3150
         ItemData        =   "FrmRestourant.frx":0CCA
         Left            =   8640
         List            =   "FrmRestourant.frx":0CCC
         Sorted          =   -1  'True
         TabIndex        =   42
         Top             =   840
         Width           =   1935
      End
      Begin VB.Label Label13 
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
         Left            =   6480
         TabIndex        =   69
         Top             =   3480
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
         Left            =   4560
         TabIndex        =   68
         Top             =   3480
         Width           =   1815
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
         Index           =   11
         Left            =   2280
         TabIndex        =   67
         Top             =   480
         Width           =   975
      End
      Begin VB.Label Label3 
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
         Left            =   240
         TabIndex        =   66
         Top             =   1320
         Width           =   1575
      End
      Begin VB.Label Label4 
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
         Left            =   5400
         TabIndex        =   65
         Top             =   1320
         Width           =   1335
      End
      Begin VB.Label Label5 
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
         Left            =   5400
         TabIndex        =   64
         Top             =   2040
         Width           =   975
      End
      Begin VB.Label Label6 
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
         Left            =   240
         TabIndex        =   63
         Top             =   480
         Width           =   1455
      End
      Begin VB.Label Label7 
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
         Left            =   240
         TabIndex        =   62
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label8 
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
         Left            =   2400
         TabIndex        =   61
         Top             =   2760
         Width           =   2055
      End
      Begin VB.Label Label9 
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
         Left            =   6960
         TabIndex        =   60
         Top             =   2760
         Width           =   1335
      End
      Begin VB.Label Label10 
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
         Left            =   4560
         TabIndex        =   59
         Top             =   2760
         Width           =   2175
      End
      Begin VB.Label Label15 
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
         Left            =   240
         TabIndex        =   58
         Top             =   3480
         Width           =   1815
      End
      Begin VB.Label Label16 
         BackStyle       =   0  'Transparent
         Caption         =   "Check Out Time"
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
         TabIndex        =   57
         Top             =   3480
         Width           =   1815
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
         Left            =   8640
         TabIndex        =   56
         Top             =   600
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Menu Items"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   360
      TabIndex        =   34
      Top             =   3360
      Width           =   6135
      Begin VB.ListBox List7 
         Appearance      =   0  'Flat
         Height          =   1980
         ItemData        =   "FrmRestourant.frx":0CCE
         Left            =   240
         List            =   "FrmRestourant.frx":0CD0
         TabIndex        =   37
         Top             =   600
         Width           =   975
      End
      Begin VB.ListBox List8 
         Appearance      =   0  'Flat
         Height          =   1980
         ItemData        =   "FrmRestourant.frx":0CD2
         Left            =   1200
         List            =   "FrmRestourant.frx":0CD4
         TabIndex        =   36
         Top             =   600
         Width           =   3735
      End
      Begin VB.ListBox List9 
         Appearance      =   0  'Flat
         Height          =   1980
         ItemData        =   "FrmRestourant.frx":0CD6
         Left            =   4920
         List            =   "FrmRestourant.frx":0CD8
         TabIndex        =   35
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   12
         Left            =   2520
         TabIndex        =   40
         Top             =   360
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
         ForeColor       =   &H00FF0000&
         Height          =   375
         Index           =   13
         Left            =   360
         TabIndex        =   39
         Top             =   360
         Width           =   975
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
         ForeColor       =   &H00FF0000&
         Height          =   255
         Index           =   14
         Left            =   5160
         TabIndex        =   38
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.TextBox Text23 
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text22 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   31
      Text            =   "0"
      Top             =   600
      Width           =   1695
   End
   Begin VB.CommandButton Command3 
      BackColor       =   &H8000000C&
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
      Left            =   1200
      Style           =   1  'Graphical
      TabIndex        =   30
      Top             =   8280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.OptionButton Option2 
      BackColor       =   &H00008080&
      Caption         =   "Hotel Customer"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   3480
      TabIndex        =   29
      Top             =   1200
      Width           =   2895
   End
   Begin VB.OptionButton Option1 
      BackColor       =   &H00008080&
      Caption         =   "Cash "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   375
      Left            =   360
      TabIndex        =   28
      Top             =   1200
      Value           =   -1  'True
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H8000000C&
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
      Left            =   4800
      Style           =   1  'Graphical
      TabIndex        =   26
      Top             =   8280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text8 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   5400
      TabIndex        =   25
      Top             =   5595
      Width           =   1095
   End
   Begin VB.TextBox Text7 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   4680
      TabIndex        =   24
      Text            =   "   Total: "
      Top             =   5595
      Width           =   735
   End
   Begin VB.TextBox Text6 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   360
      TabIndex        =   23
      Top             =   5595
      Width           =   4335
   End
   Begin VB.ListBox List5 
      Appearance      =   0  'Flat
      Height          =   1785
      ItemData        =   "FrmRestourant.frx":0CDA
      Left            =   5400
      List            =   "FrmRestourant.frx":0CDC
      TabIndex        =   21
      Top             =   3840
      Width           =   1095
   End
   Begin VB.ListBox List4 
      Appearance      =   0  'Flat
      Height          =   1785
      ItemData        =   "FrmRestourant.frx":0CDE
      Left            =   4680
      List            =   "FrmRestourant.frx":0CE0
      TabIndex        =   20
      Top             =   3840
      Width           =   735
   End
   Begin VB.ListBox List3 
      Appearance      =   0  'Flat
      Height          =   1785
      ItemData        =   "FrmRestourant.frx":0CE2
      Left            =   4080
      List            =   "FrmRestourant.frx":0CE4
      TabIndex        =   19
      Top             =   3840
      Width           =   615
   End
   Begin VB.ListBox List2 
      Appearance      =   0  'Flat
      Height          =   1785
      ItemData        =   "FrmRestourant.frx":0CE6
      Left            =   1200
      List            =   "FrmRestourant.frx":0CE8
      TabIndex        =   18
      Top             =   3840
      Width           =   2895
   End
   Begin VB.ListBox List1 
      Appearance      =   0  'Flat
      Height          =   1785
      ItemData        =   "FrmRestourant.frx":0CEA
      Left            =   360
      List            =   "FrmRestourant.frx":0CEC
      TabIndex        =   17
      Top             =   3840
      Width           =   855
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H8000000C&
      Caption         =   "OK"
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
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   8280
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   1560
      TabIndex        =   4
      Top             =   2760
      Width           =   1695
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   4680
      TabIndex        =   3
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   1560
      TabIndex        =   2
      Top             =   2280
      Width           =   1695
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   4680
      TabIndex        =   1
      Top             =   1800
      Width           =   1695
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   1560
      TabIndex        =   0
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0080FFFF&
      BorderWidth     =   3
      Height          =   7215
      Left            =   0
      Top             =   0
      Width           =   11535
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Sales Man"
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
      Index           =   16
      Left            =   3480
      TabIndex        =   33
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Bill No."
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
      Index           =   15
      Left            =   3720
      TabIndex        =   32
      Top             =   720
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Restaurant"
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
      Index           =   10
      Left            =   480
      TabIndex        =   27
      Top             =   480
      Width           =   2535
   End
   Begin VB.Label Label2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      TabIndex        =   22
      Top             =   3480
      Width           =   6135
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty."
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
      Index           =   9
      Left            =   4200
      TabIndex        =   16
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Index           =   8
      Left            =   1920
      TabIndex        =   15
      Top             =   3600
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Left            =   5760
      TabIndex        =   14
      Top             =   3600
      Width           =   975
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
      Height          =   375
      Index           =   6
      Left            =   4800
      TabIndex        =   13
      Top             =   3600
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
      Height          =   375
      Index           =   5
      Left            =   480
      TabIndex        =   12
      Top             =   3600
      Width           =   735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Total"
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
      Index           =   4
      Left            =   360
      TabIndex        =   10
      Top             =   2760
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Qty."
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
      Index           =   3
      Left            =   360
      TabIndex        =   9
      Top             =   2280
      Width           =   975
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
      Height          =   375
      Index           =   2
      Left            =   3480
      TabIndex        =   8
      Top             =   2280
      Width           =   975
   End
   Begin VB.Label Label1 
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
      Height          =   375
      Index           =   1
      Left            =   3480
      TabIndex        =   7
      Top             =   1800
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
      Height          =   375
      Index           =   0
      Left            =   360
      TabIndex        =   6
      Top             =   1800
      Width           =   975
   End
End
Attribute VB_Name = "FrmRestourant"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyDb As Database, MyRs As Recordset    ' for items dealing
Dim MyDb2 As Database, MyRs2 As Recordset  ' for item menu
Dim MyDb3 As Database, MyRs3 As Recordset  ' for sales record
Dim FS, SL
Const Reading = 1, Writing = 2

Private Sub Command1_Click()
    
     Set MyDb3 = DBEngine.Workspaces(0).OpenDatabase(RestoPath & "\RESTO.MDB")
     Set MyRs3 = MyDb3.OpenRecordset("RESTO", dbOpenDynaset)
     If Option1.Value = True Then
     MyRs3.AddNew
        MyRs3!ITEMNO = Text1.Text
        MyRs3!ITEMNAME = Text2.Text
        MyRs3!Rate = Text4.Text
        MyRs3!QTY = Text3.Text
        MyRs3!AM0UNT = Text5.Text
        MyRs3!BILLNO = Text22.Text
        MyRs3!Month = Format(Date, "MM")
        MyRs3!Date = Format(Date, "DD")
        MyRs3!Year = Format(Date, "YY")
        MyRs3!PRINT = "NOTDONE"
        MyRs3!PRINTDUPLICATE = "NOTDONE"
      MyRs3.Update
     ElseIf Option2.Value = True Then
      MyRs3.AddNew
        MyRs3!CUSTOMERNO = Text9.Text
        MyRs3!CUSTOMERNAME = Text11.Text
        MyRs3!ROOMNO = Text10.Text
        MyRs3!ADDRESS = Text13.Text
        MyRs3!ITEMNO = Text1.Text
        MyRs3!ITEMNAME = Text2.Text
        MyRs3!Rate = Text4.Text
        MyRs3!QTY = Text3.Text
        MyRs3!AM0UNT = Text5.Text
        MyRs3!BILLNO = Text22.Text
        MyRs3!Month = Format(Date, "MM")
        MyRs3!Date = Format(Date, "DD")
        MyRs3!Year = Format(Date, "YY")
        MyRs3!PRINT = "NOTDONE"
        MyRs3!PRINTDUPLICATE = "NOTDONE"
      MyRs3.Update
      MyRs.AddNew
        MyRs!SL = Text9.Text
        MyRs!Name = Text11.Text
        MyRs!ADDRESS = Text13.Text
        MyRs!ARRIVAL = Text17.Text
        MyRs!ROOMNO = Text10.Text
        
        MyRs!restitem = Text2.Text
        MyRs!itemprice = Text5.Text
        MyRs!RESTDATE = Format(Date, "DD/MM/YY")
        MyRs!RESTTIME = Time
    MyRs.Update
   End If
   StockCal
     Text1.Text = ""
     Text2.Text = ""
     Text3.Text = ""
     Text4.Text = ""
     Text5.Text = ""
        
End Sub

Private Sub Command2_Click()
quest = MsgBox("You are about to close. Do you want to keep the same no of bill. If no then the bill no will be change.", vbYesNo + vbInformation)
If quest = vbYes Then
Unload Me
Exit Sub
Else
Open MainPath & "\BILLNO.TXT" For Output As #1
Print #1, Text22.Text
Close #1
Unload Me
End If

End Sub

Private Sub Command3_Click()
Open MainPath & "\BILLNO.TXT" For Output As #1
Print #1, Text22.Text
Close #1
FrmPrintRest.Text4.Text = Text22.Text
FrmPrintRest.Show
Unload Me

End Sub

Private Sub Command4_Click()
List6_DblClick
End Sub

Private Sub Form_Load()
Me.Width = 6960
Me.Left = MDIForm1.Width / 2 - Me.Width / 2 - 300
Me.Top = MDIForm1.Height / 2 - Me.Height / 2 - 500
Shape1.Width = Me.Width - 50
On Error Resume Next

Set MyDb2 = DBEngine.Workspaces(0).OpenDatabase(ItemsPath & "\ITEMS.MDB")
Set MyRs2 = MyDb2.OpenRecordset("ITEMS", dbOpenDynaset)
   Do While Not MyRs2.EOF
      List7.AddItem MyRs2!ITEMNO
      List8.AddItem MyRs2!ITEMNAME
      List9.AddItem MyRs2!Rate
    MyRs2.MoveNext
    Loop
MyDb2.Close
Set FS = CreateObject("Scripting.FileSystemObject")
Set SL = FS.OpenTextFile(MainPath & "\BILLNO.TXT", Reading)
Text22.Text = SL.READALL
SL.Close
Text22.Text = Val(Text22.Text) + 1
 
SQL = "select * from CUSTOMER where ID='" & "YES" & "'"
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(CustomerPath & "\CUSTOMER.MDB")
Set MyRs = MyDb.OpenRecordset(SQL, dbOpenDynaset)

If MyRs.RecordCount <= 0 Then
Exit Sub
Else
   Do While Not MyRs.EOF
      List6.AddItem MyRs!SL
    MyRs.MoveNext
    Loop
End If


MyDb.Close

End Sub


Private Sub List6_DblClick()
On Error Resume Next
   Text9.Text = ""
   Text13.Text = ""
   Text11.Text = ""
   Text14.Text = ""
   Text15.Text = ""
   Text16.Text = ""
   Text17.Text = ""
   Text18.Text = ""
   Text19.Text = ""
   Text10.Text = ""
SQL = "select * from CUSTOMER where SL='" & List6.Text & "'"
Set MyDb = DBEngine.Workspaces(0).OpenDatabase(CustomerPath & "\CUSTOMER.MDB")
Set MyRs = MyDb.OpenRecordset(SQL, dbOpenDynaset)
   Text9.Text = MyRs!SL
   Text13.Text = MyRs!ADDRESS
   Text11.Text = MyRs!Name
   Text14.Text = MyRs!TEL
   Text15.Text = MyRs!EMAIL
   Text16.Text = MyRs!REGDATE
   Text17.Text = MyRs!ARRIVAL
   Text18.Text = MyRs!TYPEOFROOM
   Text19.Text = MyRs!NOOFDAYS
   Text10.Text = MyRs!ROOMNO
  Frame1.Visible = False
  Me.Width = 6960
  Me.Left = MDIForm1.Width / 2 - Me.Width / 2 - 300
  Me.Top = MDIForm1.Height / 2 - Me.Height / 2 - 500
  Shape1.Width = Me.Width - 50
  Text1.SetFocus
End Sub

Private Sub List7_Click()
SQL = "select * from ITEMS where ITEMNO='" & List7.Text & "'"
Set MyDb2 = DBEngine.Workspaces(0).OpenDatabase(ItemsPath & "\ITEMS.MDB")
Set MyRs2 = MyDb2.OpenRecordset(SQL, dbOpenDynaset)
  If MyRs2!closingstock <= 0 Then
  MsgBox "This Stock is Empty."
  Exit Sub
  End If
      Text1.Text = MyRs2!ITEMNO
      Text2.Text = MyRs2!ITEMNAME
      Text4.Text = MyRs2!Rate
List8.ListIndex = List7.ListIndex
List9.ListIndex = List7.ListIndex

End Sub

Private Sub StockCal()
  MyRs2.Edit
     MyRs2!closingstock = Val(MyRs2!closingstock) - Val(Text3.Text)
  MyRs2.Update
  MyRs2.Edit
    MyRs2!sold = Val(MyRs2!openingstock) - Val(MyRs2!closingstock)
  MyRs2.Update
End Sub

Private Sub List7_DblClick()
Text3.SetFocus
End Sub

Private Sub List7_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then

  Text3.SetFocus
 End If
 
End Sub

Private Sub Option2_Click()
If Option2.Value = True Then
Frame1.Visible = True
Me.Width = 11700
Me.Left = MDIForm1.Width / 2 - Me.Width / 2 - 300
Me.Top = MDIForm1.Height / 2 - Me.Height / 2 - 500
Shape1.Width = Me.Width - 50

List6.SetFocus
'Frame2.Visible = False

End If

End Sub

Private Sub Text1_Change()
'On Error Resume Next
'Frame2.Visible = True
'List7.SetFocus
'List7.ListIndex = 0

End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    List7.SetFocus
    'Text2.SetFocus
 End If
End Sub
Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text3.SetFocus
 End If
End Sub

Private Sub Text23_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = 13 Then
  Command1_Click
  If Text8.Text = "" Then
    Text8.Text = Text5.Text
    Else
    Text8.Text = Val(Text8.Text) + Val(Text5.Text)
    End If
    List1.AddItem Text1.Text
    List2.AddItem Text2.Text
    List3.AddItem Text3.Text
    List4.AddItem Text4.Text
    List5.AddItem Text5.Text
    Text1 = ""
    Text2 = ""
    Text3 = ""
    Text4 = ""
    Text5 = ""
    Text1.SetFocus
    
   
 End If
    
End Sub

Private Sub Text3_Change()
Text5.Text = Val(Text4.Text) * Val(Text3.Text)
If MyRs2!closingstock <= Val(Text3.Text) Then
  MsgBox "Present stock = " & MyRs2!closingstock
    Exit Sub
  End If
End Sub

Private Sub Text3_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
  Text23.SetFocus
 End If
End Sub
Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text5.SetFocus
 End If
End Sub
Private Sub Text5_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = 13 Then
    Text23.SetFocus
    End If
End Sub

Private Sub USStyle1_Click()
Command3_Click
End Sub

Private Sub USStyle2_Click()
Command1_Click
End Sub

Private Sub USStyle3_Click()
Command2_Click
End Sub
