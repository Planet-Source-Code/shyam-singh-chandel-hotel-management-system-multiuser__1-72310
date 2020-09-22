VERSION 5.00
Object = "{97824D60-58A5-4D7C-A442-3412CD2787CD}#2.0#0"; "USSTYLE.OCX"
Begin VB.Form FrmCheckOut 
   BackColor       =   &H00008080&
   BorderStyle     =   0  'None
   Caption         =   "Check Out Customers"
   ClientHeight    =   8445
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12285
   Icon            =   "FrmCheckOut.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   8445
   ScaleWidth      =   12285
   ShowInTaskbar   =   0   'False
   Begin Project1.USStyle USStyle3 
      Height          =   495
      Left            =   9240
      TabIndex        =   54
      Top             =   7560
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
   Begin Project1.USStyle USStyle2 
      Height          =   495
      Left            =   9240
      TabIndex        =   53
      Top             =   6960
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
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
      Left            =   9240
      TabIndex        =   52
      Top             =   6360
      Width           =   2535
      _ExtentX        =   4471
      _ExtentY        =   873
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Caption         =   "Done"
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
      BackColor       =   &H00808080&
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
      Height          =   495
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   51
      Top             =   6720
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H0080C0FF&
      Caption         =   "Customer No. List"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   5415
      Left            =   9240
      TabIndex        =   45
      Top             =   720
      Width           =   2535
      Begin VB.ListBox List2 
         Appearance      =   0  'Flat
         Height          =   4905
         ItemData        =   "FrmCheckOut.frx":0CCA
         Left            =   240
         List            =   "FrmCheckOut.frx":0CCC
         Sorted          =   -1  'True
         TabIndex        =   46
         Top             =   360
         Width           =   2055
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H0080C0FF&
      Caption         =   "Fooding"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   3615
      Left            =   480
      TabIndex        =   31
      Top             =   4440
      Width           =   8655
      Begin VB.TextBox Text22 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6240
         TabIndex        =   61
         Top             =   2640
         Width           =   2055
      End
      Begin VB.TextBox Text21 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6240
         TabIndex        =   59
         Top             =   2280
         Width           =   2055
      End
      Begin VB.TextBox Text20 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6240
         TabIndex        =   57
         Top             =   1920
         Width           =   2055
      End
      Begin VB.TextBox Text19 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6240
         TabIndex        =   55
         Top             =   1560
         Width           =   2055
      End
      Begin VB.TextBox Text17 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6240
         TabIndex        =   43
         Top             =   840
         Width           =   2055
      End
      Begin VB.TextBox Text16 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6240
         TabIndex        =   38
         Top             =   3000
         Width           =   2055
      End
      Begin VB.ListBox List1 
         Appearance      =   0  'Flat
         Height          =   2955
         ItemData        =   "FrmCheckOut.frx":0CCE
         Left            =   240
         List            =   "FrmCheckOut.frx":0CD0
         TabIndex        =   35
         Top             =   480
         Width           =   3135
      End
      Begin VB.ListBox List3 
         Appearance      =   0  'Flat
         Height          =   2955
         ItemData        =   "FrmCheckOut.frx":0CD2
         Left            =   3360
         List            =   "FrmCheckOut.frx":0CD4
         TabIndex        =   34
         Top             =   480
         Width           =   1215
      End
      Begin VB.TextBox Text11 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6240
         TabIndex        =   33
         Top             =   1200
         Width           =   2055
      End
      Begin VB.TextBox Text13 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   6240
         TabIndex        =   32
         Top             =   480
         Width           =   2055
      End
      Begin VB.Label Label28 
         BackStyle       =   0  'Transparent
         Caption         =   "Bar Amt."
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
         TabIndex        =   62
         Top             =   2640
         Width           =   1215
      End
      Begin VB.Label Label27 
         BackStyle       =   0  'Transparent
         Caption         =   "Tel. Bill"
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
         TabIndex        =   60
         Top             =   2280
         Width           =   1455
      End
      Begin VB.Label Label26 
         BackStyle       =   0  'Transparent
         Caption         =   "Room Service"
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
         TabIndex        =   58
         Top             =   1920
         Width           =   1455
      End
      Begin VB.Label Label25 
         BackStyle       =   0  'Transparent
         Caption         =   "Tax"
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
         TabIndex        =   56
         Top             =   1560
         Width           =   1215
      End
      Begin VB.Label Label22 
         BackStyle       =   0  'Transparent
         Caption         =   "Restaurant"
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
         TabIndex        =   44
         Top             =   1200
         Width           =   1215
      End
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "Price"
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
         Left            =   3360
         TabIndex        =   42
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "Menu"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label Label19 
         BackColor       =   &H0080C0FF&
         Caption         =   "Total Amounts"
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
         Left            =   6120
         TabIndex        =   40
         Top             =   0
         Width           =   2175
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "Total Amt."
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
         TabIndex        =   39
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label Label11 
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
         Left            =   4800
         TabIndex        =   37
         Top             =   840
         Width           =   1215
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
         Left            =   4800
         TabIndex        =   36
         Top             =   480
         Width           =   1335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H0080C0FF&
      Caption         =   "Lodging"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   3735
      Left            =   480
      TabIndex        =   4
      Top             =   720
      Width           =   8655
      Begin VB.TextBox Text18 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   7200
         TabIndex        =   49
         Text            =   "0"
         Top             =   3240
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.ComboBox Combo1 
         Appearance      =   0  'Flat
         Height          =   315
         ItemData        =   "FrmCheckOut.frx":0CD6
         Left            =   5760
         List            =   "FrmCheckOut.frx":0CE3
         TabIndex        =   48
         Top             =   3240
         Width           =   1335
      End
      Begin VB.TextBox Text12 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4080
         TabIndex        =   17
         Top             =   3240
         Width           =   1575
      End
      Begin VB.TextBox Text10 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   7320
         TabIndex        =   16
         Top             =   2640
         Width           =   1095
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   240
         TabIndex        =   15
         Top             =   600
         Width           =   1935
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2280
         TabIndex        =   14
         Top             =   600
         Width           =   6135
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   1095
         Left            =   240
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   13
         Top             =   1250
         Width           =   5055
      End
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5400
         TabIndex        =   12
         Top             =   1250
         Width           =   3015
      End
      Begin VB.TextBox Text5 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5400
         TabIndex        =   11
         Top             =   1920
         Width           =   3015
      End
      Begin VB.TextBox Text6 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox Text7 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2160
         TabIndex        =   9
         Top             =   2640
         Width           =   1815
      End
      Begin VB.TextBox Text8 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   4080
         TabIndex        =   8
         Top             =   2640
         Width           =   1575
      End
      Begin VB.TextBox Text9 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   5760
         TabIndex        =   7
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox Text14 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   240
         TabIndex        =   6
         Top             =   3240
         Width           =   1815
      End
      Begin VB.TextBox Text15 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   2160
         TabIndex        =   5
         Top             =   3240
         Width           =   1815
      End
      Begin VB.Label Label24 
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque No."
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   7200
         TabIndex        =   50
         Top             =   3000
         Visible         =   0   'False
         Width           =   1335
      End
      Begin VB.Label Label23 
         BackStyle       =   0  'Transparent
         Caption         =   "Payment By"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   5760
         TabIndex        =   47
         Top             =   3000
         Width           =   1455
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
         Height          =   255
         Left            =   7320
         TabIndex        =   30
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label12 
         BackStyle       =   0  'Transparent
         Caption         =   "Room Charges"
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
         Left            =   4080
         TabIndex        =   29
         Top             =   3000
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
         Left            =   2280
         TabIndex        =   28
         Top             =   360
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
         Left            =   240
         TabIndex        =   27
         Top             =   1000
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
         Left            =   5400
         TabIndex        =   26
         Top             =   1000
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
         Left            =   5400
         TabIndex        =   25
         Top             =   1680
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
         Left            =   240
         TabIndex        =   24
         Top             =   360
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
         Left            =   240
         TabIndex        =   23
         Top             =   2400
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
         Left            =   2160
         TabIndex        =   22
         Top             =   2400
         Width           =   2055
      End
      Begin VB.Label Label8 
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
         Left            =   5760
         TabIndex        =   21
         Top             =   2400
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
         Left            =   4080
         TabIndex        =   20
         Top             =   2400
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
         TabIndex        =   19
         Top             =   3000
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
         Left            =   2160
         TabIndex        =   18
         Top             =   3000
         Width           =   1815
      End
   End
   Begin VB.Timer Timer1 
      Interval        =   500
      Left            =   5040
      Top             =   7200
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
      Height          =   495
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7320
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.CommandButton Command1 
      BackColor       =   &H00808080&
      Caption         =   "Done"
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
      Height          =   495
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   6120
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   3
      Height          =   8415
      Left            =   20
      Top             =   20
      Width           =   12255
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Customer No. List"
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
      Left            =   12600
      TabIndex        =   3
      Top             =   360
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Check Out Customers"
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
      Left            =   480
      TabIndex        =   2
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "FrmCheckOut"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyDb As Database, MyRs As Recordset
Dim MyDb1 As Database, MyRs1 As Recordset
Dim Adv As Single, TOT As Single
Dim Db As Database, Rs As Recordset


Private Sub Combo1_Click()
If Combo1.Text = "CHEQUE" Then
Text18.Visible = True
Label24.Caption = "Cheque No."
Label24.Visible = True
ElseIf Combo1.Text = "D.D." Then
Text18.Visible = True
Label24.Caption = "DD No."
Label24.Visible = True
Else
Text18.Visible = False
Label24.Visible = False
End If
End Sub

Private Sub Command1_Click()
On Error Resume Next
If Combo1.Text = "" Then
MsgBox "Please select the Payment By"
Exit Sub
End If
If Text11.Text = "" Then
Text11.Text = 0
End If
If Text17.Text = "" Then
Text17.Text = 0
End If

Set MyDb1 = DBEngine.Workspaces(0).OpenDatabase(PrintPath & "\PRINT.MDB")
Set MyRs1 = MyDb1.OpenRecordset("PRINTDB", dbOpenDynaset)
     MyRs1.AddNew
        MyRs1!ID = Text1.Text
        MyRs1!Name = Text2.Text
        MyRs1!ADDRESS = Text3.Text
        MyRs1!lodging = Text13.Text
        MyRs1!FOODING = Text11.Text
        MyRs1!ADVANCE = Text17.Text
        MyRs1!ROOMNO = Text10.Text
        MyRs1!TYPEOFROOM = Text8.Text
        MyRs1!checkindate = Text7.Text
        MyRs1!checkoutdate = Text14.Text
        MyRs1!CHECKOUTTIME = Text15.Text
        MyRs1!NETAMOUNT = Text16.Text
        MyRs1!printstatus = "NOTDONE"
        MyRs1!roomcharges = Text12.Text
        MyRs1!NOOFDAYS = Text9.Text
       MyRs1.Update
       
                         
       SQL = "select * from CUSTOMER where SL='" & List2.Text & "'"
       Set MyDb = DBEngine.Workspaces(0).OpenDatabase(CustomerPath & "\CUSTOMER.MDB")
       Set MyRs = MyDb.OpenRecordset(SQL, dbOpenDynaset)
       Do While Not MyRs.EOF
            MyRs.Edit
              MyRs!BILLINGTIME = Format(Now, "HH:MM AM/PM")
              MyRs!BILLAMOUNT = TOT
              MyRs!CHECKOUTSTATUS = "DONE"
              MyRs!BILLBALANCE = Text16.Text
              MyRs!BILLPAYMENTBY = Combo1.Text
              MyRs!CH_DD_NO = Text18.Text
            MyRs.Update
            MyRs.MoveNext
       Loop
        SQL = "select * from ROOMS where ROOMNO='" & Text10.Text & "'"
            Set Db = DBEngine.Workspaces(0).OpenDatabase(RoomsPath & "\ROOMS.MDB")
            Set Rs = Db.OpenRecordset(SQL, dbOpenDynaset)
            
            Rs.Edit
            Rs!Status = "EMPTY"
            Rs.Update
            Rs.Close
            Db.Close
       MsgBox "Check Out Bill Has been Created"
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
     Text16.Text = ""
     Text17.Text = ""
     List2.Clear
     Form_Load
End Sub

Private Sub Command2_Click()
FrmPrintBill.Show
Unload Me

End Sub

Private Sub Command4_Click()
Unload Me

End Sub

Private Sub Form_Load()
Me.Left = MDIForm1.Width / 2 - Me.Width / 2 - 300
Me.Top = MDIForm1.Height / 2 - Me.Height / 2 - 500

SQL = "select distinct sl from CUSTOMER where CHECKOUTSTATUS='" & "NOTDONE" & "'"
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
'MyDb.Close
Text14.Text = Format(Now, "DD/MM/yyyy")
Text15.Text = Format(Now, "hh:mm:ss AM/PM")
End Sub

Private Sub List2_Click()
On Error Resume Next
Dim RC
     List1.Clear
     List3.Clear
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
   Text14.Text = MyRs!checkoutdate
   Text12.Text = MyRs!roomcharges
   Text17.Text = MyRs!ADVANCE
   Text15.Text = MyRs!CHECKOUTTIME
   MyRs.MoveFirst
 Do While Not MyRs.EOF
      List1.AddItem MyRs!restitem
      RC = MyRs!itemprice
      List3.AddItem RC
      Text11.Text = Val(Text11.Text) + RC
     
    MyRs.MoveNext
    Loop
    

Text13.Text = Val(Text9.Text) * Val(Text12.Text)
Adv = Val(Text13.Text) - Val(Text17.Text)
Text16.Text = Adv + Val(Text11.Text)
TOT = Val(Text13.Text) + Val(Text11.Text)
Command1.Enabled = True
Command2.Enabled = True
End Sub

Private Sub Text11_Change()
Dim Adv As Single
Adv = Val(Text13.Text) - Val(Text17.Text)
Text16.Text = Adv + Val(Text11.Text)

End Sub

Private Sub USStyle1_Click()
Command1_Click
End Sub

Private Sub USStyle2_Click()
Command2_Click
End Sub

Private Sub USStyle3_Click()
Unload Me
End Sub
