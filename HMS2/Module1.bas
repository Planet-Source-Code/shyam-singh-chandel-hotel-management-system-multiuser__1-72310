Attribute VB_Name = "Module1"
Option Explicit
Option Compare Text
Global RunSt As String
Global MainPath As String
Global RestoPath As String
Global StaffPath As String
Global CustomerPath As String
Global ItemsPath As String
Global PrintPath As String
Global RoomsPath As String
Global UserPath As String
Global SCRTIME As Integer
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long

#If Win32 Then
    Public Const CB_FINDSTRING = &H14C
    Public Const CB_FINDSTRINGEXACT = &H158
    Public Const LB_FINDSTRING = &H18F
    Public Const LB_FINDSTRINGEXACT = &H1A2
#Else
    Public Const WM_USER = &H400
    Public Const CB_FINDSTRING = WM_USER + 12
    Public Const CB_FINDSTRINGEXACT = WM_USER + 24
    Public Const LB_FINDSTRING = WM_USER + 16
    Public Const LB_FINDSTRINGEXACT = WM_USER + 35
#End If

Public Function FindFirstMatch(ByVal ctlSearch As Control, ByVal SearchString As String, ByVal FirstRow As Integer, ByVal Exact As Boolean) As Integer

#If Win32 Then
    Dim Index As Long
#Else
    Dim Index As Integer
#End If

On Error Resume Next
If TypeOf ctlSearch Is ComboBox Then
    If Exact Then
        Index = SendMessage(ctlSearch.hwnd, CB_FINDSTRINGEXACT, FirstRow, ByVal SearchString)
    Else
        Index = SendMessage(ctlSearch.hwnd, CB_FINDSTRING, FirstRow, ByVal SearchString)
    End If
ElseIf TypeOf ctlSearch Is ListBox Then
    If Exact Then
        Index = SendMessage(ctlSearch.hwnd, LB_FINDSTRINGEXACT, FirstRow, ByVal SearchString)
    Else
        Index = SendMessage(ctlSearch.hwnd, LB_FINDSTRING, FirstRow, ByVal SearchString)
    End If
End If

FindFirstMatch = Index

End Function


Public Function CUSTOMER(DBFullPath As String) As Boolean
Dim Db As Database
Dim TD  As TableDef

Dim f As Field

On Error GoTo ErrorHandler
' Return reference to current database.
Set Db = DBEngine.CreateDatabase(DBFullPath, dbLangGeneral)
' Create new TableDef object.
Set TD = Db.CreateTableDef("CUSTOMER")
' Create new Field object.

Set f = TD.CreateField("ID", dbText)
TD.Fields.Append f
Set f = TD.CreateField("SL", dbText)
TD.Fields.Append f
Set f = TD.CreateField("NAME", dbText)
TD.Fields.Append f
Set f = TD.CreateField("ADDRESS", dbText)
TD.Fields.Append f
Set f = TD.CreateField("Tel", dbText)
TD.Fields.Append f
Set f = TD.CreateField("Email", dbText)
TD.Fields.Append f
Set f = TD.CreateField("CHECKOUTSTATUS", dbText)
TD.Fields.Append f
Set f = TD.CreateField("ARRIVAL", dbText)
TD.Fields.Append f
Set f = TD.CreateField("REGEXPIRY", dbText)
TD.Fields.Append f
Set f = TD.CreateField("REGDATE", dbText)
TD.Fields.Append f
Set f = TD.CreateField("ADVANCE", dbText)
TD.Fields.Append f
Set f = TD.CreateField("BALANCE", dbText)
TD.Fields.Append f
Set f = TD.CreateField("CHECKINDATE", dbText)
TD.Fields.Append f
Set f = TD.CreateField("CHECKINTIME", dbText)
TD.Fields.Append f
Set f = TD.CreateField("CHECKOUTDATE", dbText)
TD.Fields.Append f
Set f = TD.CreateField("CHECKOUTTIME", dbText)
TD.Fields.Append f
Set f = TD.CreateField("NOOFDAYS", dbText)
TD.Fields.Append f
Set f = TD.CreateField("ITEMNO", dbText)
TD.Fields.Append f
Set f = TD.CreateField("RESTITEM", dbText)
TD.Fields.Append f
Set f = TD.CreateField("ITEMPRICE", dbText)
TD.Fields.Append f
Set f = TD.CreateField("RESTDATE", dbText)
TD.Fields.Append f
Set f = TD.CreateField("RESTTIME", dbText)
TD.Fields.Append f
Set f = TD.CreateField("ROOMNO", dbText)
TD.Fields.Append f
Set f = TD.CreateField("TYPEOFROOM", dbText)
TD.Fields.Append f
Set f = TD.CreateField("ROOMCHARGES", dbText)
TD.Fields.Append f
Set f = TD.CreateField("ANYEXTRA", dbText)
TD.Fields.Append f
Set f = TD.CreateField("NOTES", dbText)
TD.Fields.Append f
Set f = TD.CreateField("BILLINGTIME", dbText)
TD.Fields.Append f
Set f = TD.CreateField("BILLAMOUNT", dbText)
TD.Fields.Append f
Set f = TD.CreateField("BILLBALANCE", dbText)
TD.Fields.Append f
Set f = TD.CreateField("BILLPAYMENTBY", dbText)
TD.Fields.Append f
Set f = TD.CreateField("CH_DD_NO", dbText)
TD.Fields.Append f
Set f = TD.CreateField("CH_DD_BANKNAME", dbText)
TD.Fields.Append f




Db.TableDefs.Append TD ''

CUSTOMER = True
ErrorHandler:
If Not Db Is Nothing Then Db.Close
End Function

Public Function STAFF(DBFullPath As String) As Boolean
Dim Db As Database
Dim TD  As TableDef

Dim f As Field

On Error GoTo ErrorHandler
' Return reference to current database.
Set Db = DBEngine.CreateDatabase(DBFullPath, dbLangGeneral)
' Create new TableDef object.
Set TD = Db.CreateTableDef("STAFF")
' Create new Field object.

Set f = TD.CreateField("ID", dbText)
TD.Fields.Append f
Set f = TD.CreateField("NAME", dbText)
TD.Fields.Append f
Set f = TD.CreateField("ADDRESS", dbMemo)
TD.Fields.Append f
Set f = TD.CreateField("Tel", dbText)
TD.Fields.Append f
Set f = TD.CreateField("Email", dbText)
TD.Fields.Append f
Set f = TD.CreateField("BASICPAY", dbText)
TD.Fields.Append f
Set f = TD.CreateField("ADVANCEPAY", dbText)
TD.Fields.Append f
Set f = TD.CreateField("BALANCEPAY", dbText)
TD.Fields.Append f
Set f = TD.CreateField("NOTES", dbText)
TD.Fields.Append f
Set f = TD.CreateField("JOININGDATE", dbText)
TD.Fields.Append f
Set f = TD.CreateField("RESIGNATIONDATE", dbText)
TD.Fields.Append f
Set f = TD.CreateField("NOOFMONTH", dbText)
TD.Fields.Append f

Db.TableDefs.Append TD ''

STAFF = True
ErrorHandler:
If Not Db Is Nothing Then Db.Close

End Function
Public Function ITEMS(DBFullPath As String) As Boolean
Dim Db As Database
Dim TD  As TableDef

Dim f As Field

On Error GoTo ErrorHandler
' Return reference to current database.
Set Db = DBEngine.CreateDatabase(DBFullPath, dbLangGeneral)
' Create new TableDef object.
Set TD = Db.CreateTableDef("ITEMS")
' Create new Field object.

Set f = TD.CreateField("ID", dbText)
TD.Fields.Append f
Set f = TD.CreateField("ITEMNO", dbText)
TD.Fields.Append f
Set f = TD.CreateField("ITEMNAME", dbText)
TD.Fields.Append f
Set f = TD.CreateField("RATE", dbText)
TD.Fields.Append f
Set f = TD.CreateField("OPENINGSTOCK", dbText)
TD.Fields.Append f
Set f = TD.CreateField("CLOSINGSTOCK", dbText)
TD.Fields.Append f
Set f = TD.CreateField("SOLD", dbText)
TD.Fields.Append f
Set f = TD.CreateField("OPENINGSTOCKDATE", dbText)
TD.Fields.Append f
Set f = TD.CreateField("NOTES", dbText)
TD.Fields.Append f
Set f = TD.CreateField("SALESMANE", dbText)
TD.Fields.Append f
Set f = TD.CreateField("STOCKENTRYPERSON", dbText)
TD.Fields.Append f

Db.TableDefs.Append TD ''

ITEMS = True
ErrorHandler:
If Not Db Is Nothing Then Db.Close

End Function

Public Function RESTO(DBFullPath As String) As Boolean
Dim Db As Database
Dim TD  As TableDef

Dim f As Field

On Error GoTo ErrorHandler
' Return reference to current database.
Set Db = DBEngine.CreateDatabase(DBFullPath, dbLangGeneral)
' Create new TableDef object.
Set TD = Db.CreateTableDef("RESTO")
' Create new Field object.

Set f = TD.CreateField("ID", dbText)
TD.Fields.Append f
Set f = TD.CreateField("ITEMNO", dbText)
TD.Fields.Append f
Set f = TD.CreateField("ITEMNAME", dbText)
TD.Fields.Append f
Set f = TD.CreateField("RATE", dbText)
TD.Fields.Append f
Set f = TD.CreateField("QTY", dbText)
TD.Fields.Append f
Set f = TD.CreateField("AM0UNT", dbText)
TD.Fields.Append f
Set f = TD.CreateField("MONTH", dbText)
TD.Fields.Append f
Set f = TD.CreateField("DATE", dbText)
TD.Fields.Append f
Set f = TD.CreateField("YEAR", dbText)
TD.Fields.Append f
Set f = TD.CreateField("SALESMANE", dbText)
TD.Fields.Append f
Set f = TD.CreateField("CUSTOMERNO", dbText)
TD.Fields.Append f
Set f = TD.CreateField("CUSTOMERNAME", dbText)
TD.Fields.Append f
Set f = TD.CreateField("ADDRESS", dbText)
TD.Fields.Append f
Set f = TD.CreateField("ROOMNO", dbText)
TD.Fields.Append f
Set f = TD.CreateField("BILLNO", dbText)
TD.Fields.Append f
Set f = TD.CreateField("PRINT", dbText)
TD.Fields.Append f
Set f = TD.CreateField("PRINTDUPLICATE", dbText)
TD.Fields.Append f

Db.TableDefs.Append TD ''

RESTO = True
ErrorHandler:
If Not Db Is Nothing Then Db.Close

End Function

Public Function PRINTDB(DBFullPath As String) As Boolean
Dim Db As Database
Dim TD  As TableDef

Dim f As Field

On Error GoTo ErrorHandler
' Return reference to current database.
Set Db = DBEngine.CreateDatabase(DBFullPath, dbLangGeneral)
' Create new TableDef object.
Set TD = Db.CreateTableDef("PRINTDB")
' Create new Field object.

Set f = TD.CreateField("ID", dbText)
TD.Fields.Append f
Set f = TD.CreateField("NAME", dbText)
TD.Fields.Append f
Set f = TD.CreateField("ADDRESS", dbText)
TD.Fields.Append f
Set f = TD.CreateField("LODGING", dbText)
TD.Fields.Append f
Set f = TD.CreateField("FOODING", dbText)
TD.Fields.Append f
Set f = TD.CreateField("TAX", dbText)
TD.Fields.Append f
Set f = TD.CreateField("NETAMOUNT", dbText)
TD.Fields.Append f
Set f = TD.CreateField("NOOFDAYS", dbText)
TD.Fields.Append f
Set f = TD.CreateField("ADVANCE", dbText)
TD.Fields.Append f
Set f = TD.CreateField("ROOMNO", dbText)
TD.Fields.Append f
Set f = TD.CreateField("TYPEOFROOM", dbText)
TD.Fields.Append f
Set f = TD.CreateField("ROOMCHARGES", dbText)
TD.Fields.Append f
Set f = TD.CreateField("SERVICE", dbText)
TD.Fields.Append f
Set f = TD.CreateField("RECMAN", dbText)
TD.Fields.Append f
Set f = TD.CreateField("CHECKINDATE", dbText)
TD.Fields.Append f
Set f = TD.CreateField("CHECKOUTDATE", dbText)
TD.Fields.Append f
Set f = TD.CreateField("CHECKOUTTIME", dbText)
TD.Fields.Append f
Set f = TD.CreateField("NOTES", dbText)
TD.Fields.Append f
Set f = TD.CreateField("BALANCE", dbText)
TD.Fields.Append f
Set f = TD.CreateField("PRINTSTATUS", dbText)
TD.Fields.Append f
Set f = TD.CreateField("BILLINGDATE", dbText)
TD.Fields.Append f
Set f = TD.CreateField("DUPLICATEBILLSTATUS", dbText)
TD.Fields.Append f
Set f = TD.CreateField("DUPLICATEBILLDATE", dbText)
TD.Fields.Append f

Db.TableDefs.Append TD ''

PRINTDB = True
ErrorHandler:
If Not Db Is Nothing Then Db.Close

End Function

Public Function USERDB(DBFullPath As String) As Boolean
Dim Db As Database
Dim TD  As TableDef

Dim f As Field

On Error GoTo ErrorHandler
' Return reference to current database.
Set Db = DBEngine.CreateDatabase(DBFullPath, dbLangGeneral)
' Create new TableDef object.
Set TD = Db.CreateTableDef("USERDB")
' Create new Field object.

Set f = TD.CreateField("ID", dbText)
TD.Fields.Append f
Set f = TD.CreateField("NAME", dbText)
TD.Fields.Append f
Set f = TD.CreateField("USERNAME", dbText)
TD.Fields.Append f
Set f = TD.CreateField("PASS", dbText)
TD.Fields.Append f


Db.TableDefs.Append TD ''

USERDB = True
ErrorHandler:
If Not Db Is Nothing Then Db.Close

End Function

Public Function ROOMS(DBFullPath As String) As Boolean
Dim Db As Database
Dim TD  As TableDef

Dim f As Field

On Error GoTo ErrorHandler
' Return reference to current database.
Set Db = DBEngine.CreateDatabase(DBFullPath, dbLangGeneral)
' Create new TableDef object.
Set TD = Db.CreateTableDef("ROOMS")
' Create new Field object.

Set f = TD.CreateField("ID", dbText)
TD.Fields.Append f
Set f = TD.CreateField("ROOMNO", dbText)
TD.Fields.Append f
Set f = TD.CreateField("TYPEOFROOM", dbText)
TD.Fields.Append f
Set f = TD.CreateField("RATE", dbText)
TD.Fields.Append f
Set f = TD.CreateField("STATUS", dbText)
TD.Fields.Append f
Set f = TD.CreateField("BOOKINGDATE", dbText)
TD.Fields.Append f
Set f = TD.CreateField("FEATURS", dbText)
TD.Fields.Append f


Db.TableDefs.Append TD ''

ROOMS = True
ErrorHandler:
If Not Db Is Nothing Then Db.Close

End Function

