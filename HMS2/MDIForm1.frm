VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H00008080&
   Caption         =   "Hotel Management System"
   ClientHeight    =   10710
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   15240
   Icon            =   "MDIForm1.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      Height          =   14535
      Left            =   0
      ScaleHeight     =   14475
      ScaleWidth      =   15180
      TabIndex        =   0
      Top             =   0
      Width           =   15240
      Begin VB.Timer Timer1 
         Enabled         =   0   'False
         Interval        =   4000
         Left            =   120
         Top             =   120
      End
      Begin VB.Image Image1 
         Height          =   14535
         Left            =   0
         Stretch         =   -1  'True
         Top             =   0
         Width           =   18975
      End
   End
   Begin VB.Menu FILE 
      Caption         =   "File"
      Begin VB.Menu MnuRestCashMemo 
         Caption         =   "Print Restaurant Cash Memo"
      End
      Begin VB.Menu a 
         Caption         =   "-"
      End
      Begin VB.Menu MnuCheckoutbill 
         Caption         =   "Print Checkout Bill"
      End
      Begin VB.Menu b 
         Caption         =   "-"
      End
      Begin VB.Menu MnuExit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu MnuAdmin 
      Caption         =   "Admin"
      Begin VB.Menu MnuUser 
         Caption         =   "User"
      End
      Begin VB.Menu MnuRoomEntry 
         Caption         =   "Room Entry"
      End
      Begin VB.Menu MnuLogout 
         Caption         =   "Logout"
      End
      Begin VB.Menu scrsetting 
         Caption         =   "Screen Saver Time Settings"
      End
   End
   Begin VB.Menu MnuView 
      Caption         =   "View"
      Begin VB.Menu roomcharges 
         Caption         =   "Room Charges"
      End
      Begin VB.Menu MnuRoomStatus 
         Caption         =   "Room Status"
      End
      Begin VB.Menu MnuReservStatus 
         Caption         =   "Reservation Status"
      End
      Begin VB.Menu MnuViewCheckIn 
         Caption         =   "Check In"
      End
      Begin VB.Menu MnuViewCheckOut 
         Caption         =   "Check Out"
      End
   End
   Begin VB.Menu MnuReception 
      Caption         =   "Reception"
      Begin VB.Menu MnuReservation 
         Caption         =   "Room Reservation   ""Add/Edit/Delete"""
      End
      Begin VB.Menu MnuConfermReser 
         Caption         =   "Confirm Reservation"
      End
      Begin VB.Menu MnuCheckOut 
         Caption         =   "Check Out"
      End
   End
   Begin VB.Menu MnuRestourant 
      Caption         =   "Restaurant"
      Begin VB.Menu MnuAddMenue 
         Caption         =   "Menu Stock Entry"
      End
      Begin VB.Menu MnuRestBilling 
         Caption         =   "Billing"
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "Help ?"
   End
   Begin VB.Menu MnuAbout 
      Caption         =   "About"
   End
   Begin VB.Menu exit 
      Caption         =   "Exit"
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim MyDb As Database, MyRs As Recordset
Dim PIC As Integer

Private Sub exit_Click()
PlayWave App.Path & "\123\DRUM CLICK SOUND.wav"
res = MsgBox("Do you want to Quit ?", vbYesNo + vbQuestion, "Want to Quit ?")
If res = vbYes Then
FrmGoodBy.Show
Unload Me
Else
Exit Sub
End If

End Sub

Private Sub MDIForm_Load()
PIC = 0
Me.Left = 0
Me.Top = 0
Picture1.Visible = False
RunStatus = "frmAbout"
MDIForm1.Caption = "Hotel Management System   ::::Not IDLE::::"
 FrmIdle.Show
End Sub

Private Sub MDIForm_Resize()
'Picture1.Width = Me.Width
'Picture1.Height = Me.Height
Image1.Width = Me.Width
Image1.Height = Me.Height
End Sub

Private Sub MDIForm_Unload(Cancel As Integer)
Cancel = True

End Sub

Private Sub MnuAbout_Click()
frmAbout.Show
PlayWave App.Path & "\123\DRUM CLICK SOUND.wav"
End Sub

Private Sub MnuAddMenue_Click()
FrmStockEntry.Show
End Sub

Private Sub MnuCheckOut_Click()
FrmCheckOut.Show
End Sub

Private Sub MnuCheckoutbill_Click()
FrmPrintBill.Show

End Sub

Private Sub MnuConfermReser_Click()
FrmConferm.Show
End Sub

Private Sub MnuExit_Click()

exit_Click

End Sub

Private Sub MnuHelp_Click()
FrmHelp.Show
End Sub

Private Sub MnuLogout_Click()
FrmLogin.Show
Me.Hide
End Sub

Private Sub MnuReservation_Click()
FrmRegister.Show
End Sub

Private Sub MnuReservStatus_Click()
FrmRegistrationStatus.Show
End Sub

Private Sub MnuRestBilling_Click()
FrmRestourant.Show
End Sub

Private Sub MnuRestCashMemo_Click()
FrmPrintRest.Show

End Sub

Private Sub MnuRoomEntry_Click()
FrmRoomEntry.Show
End Sub

Private Sub MnuRoomStatus_Click()
FrmRoomStatus.Show
End Sub

Private Sub MnuUser_Click()
FrmCreatUser.Show
End Sub

Private Sub MnuViewCheckIn_Click()
FrmCheckInShow.Show
End Sub

Private Sub MnuViewCheckOut_Click()
FrmCheckOutView.Show

End Sub

Private Sub roomcharges_Click()
FrmRoomInfo.Show
End Sub

Private Sub scrsetting_Click()
Dim ins
ins = InputBox("Please Enter the time in second", "Screen Saver Settings")
Call SaveSetting("USHMS", "SETTINGS", "0US009", ins)
MsgBox "settings has been saved please restart the application now"
End

End Sub

Private Sub Timer1_Timer()
PIC = PIC + 1
Image1.Picture = LoadPicture(App.Path & "\SCR\" & PIC & ".JPG")
'PlayWave App.Path & "\123\GOLF GROUND MUSIC.wav"
If PIC >= 6 Then
PIC = 0
End If

End Sub

