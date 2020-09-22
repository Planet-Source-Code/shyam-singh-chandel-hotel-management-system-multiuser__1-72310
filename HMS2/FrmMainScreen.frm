VERSION 5.00
Begin VB.Form FrmMainScreen 
   Caption         =   "Hotel Management System"
   ClientHeight    =   6495
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   9510
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Picture         =   "FrmMainScreen.frx":0000
   ScaleHeight     =   6495
   ScaleWidth      =   9510
   Begin VB.Menu MnuAdmin 
      Caption         =   "Admin"
      Begin VB.Menu MnuUser 
         Caption         =   "User"
      End
      Begin VB.Menu MnuDutyChart 
         Caption         =   "Duty Chart"
      End
      Begin VB.Menu MnuStaff 
         Caption         =   "Staff"
      End
      Begin VB.Menu MnuStaffPay 
         Caption         =   "Payroll"
      End
   End
   Begin VB.Menu MnuView 
      Caption         =   "View"
      Begin VB.Menu MnuroomStatus 
         Caption         =   "Room Status"
      End
      Begin VB.Menu MnuClassify 
         Caption         =   "Classi fied Rooms"
      End
      Begin VB.Menu MnuViewCustAdvance 
         Caption         =   "Customer Advance"
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
         Caption         =   "Room Reservation"
      End
      Begin VB.Menu MnuConfermReser 
         Caption         =   "Conferm/Cancelation Reservation"
      End
      Begin VB.Menu MnuCheckIn 
         Caption         =   "Check In"
      End
      Begin VB.Menu MnuCheckOut 
         Caption         =   "Check Out"
      End
      Begin VB.Menu MnuCustomerAdvance 
         Caption         =   "Customer Advance"
      End
      Begin VB.Menu MnuBilling 
         Caption         =   "Billing"
      End
   End
   Begin VB.Menu MnuRestourant 
      Caption         =   "Restourent"
      Begin VB.Menu MnuAddMenue 
         Caption         =   "Add Menue"
      End
      Begin VB.Menu MnuMenueRates 
         Caption         =   "View Mnue Rates"
      End
      Begin VB.Menu MnuRestBillong 
         Caption         =   "Billing"
      End
      Begin VB.Menu MnuBar 
         Caption         =   "Bar"
      End
      Begin VB.Menu MnuPagPrises 
         Caption         =   "Pag Prices"
      End
      Begin VB.Menu MnuBarBilling 
         Caption         =   "Billing"
      End
   End
   Begin VB.Menu MnuHelp 
      Caption         =   "Help?"
   End
End
Attribute VB_Name = "FrmMainScreen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Left = 10000
End Sub

Private Sub MnuConfermReser_Click()
FrmConferm.Show
End Sub

Private Sub MnuReservation_Click()
FrmRegister.Show
End Sub

Private Sub MnuRestBillong_Click()
FrmRestourant.Show
End Sub
