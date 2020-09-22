Attribute VB_Name = "mIDLE"

Option Explicit

'---------------------------------------------------------------------------------------
' CONSTANT
'---------------------------------------------------------------------------------------
Private Const ERROR As Long = 0
Private Const WH_KEYBOARD As Long = 2
Private Const WH_MOUSE As Long = 7

'---------------------------------------------------------------------------------------
' TYPES
'---------------------------------------------------------------------------------------
Private Type LASTINPUTINFO
    lSize As Long
    lTime As Long
End Type

'The OSVERSIONINFO data structure contains operating system version information.
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

'---------------------------------------------------------------------------------------
' APIS
'---------------------------------------------------------------------------------------
Private Declare Function SetTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetLastInputInfo Lib "user32.dll" (plii As LASTINPUTINFO) As Long
Private Declare Function GetTickCount Lib "kernel32" () As Long
Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function CallNextHookEx Lib "user32" (ByVal hHook As Long, ByVal ncode As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function SetWindowsHookEx Lib "user32" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32" (ByVal hHook As Long) As Long

'---------------------------------------------------------------------------------------
' MEMBER VARIABLES
'---------------------------------------------------------------------------------------
Private m_lhWnd As Long
Private m_lIDLETime As Long
Private m_bInIDLE As Boolean
Private m_bGetLastInputInfo As Boolean
Private m_lKeyHook As Long
Private m_lMouseHook As Long
Private m_lTime As Long
'
'---------------------------------------------------------------------------------------
' Procedure : IDLE
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Let IDLE(lTime As Long)
    m_lIDLETime = lTime
End Property
'
'---------------------------------------------------------------------------------------
' Procedure : IDLE
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get IDLE() As Long
    IDLE = m_lIDLETime
End Property
'
'---------------------------------------------------------------------------------------
' Procedure : InIDLE
' Purpose   :
'---------------------------------------------------------------------------------------
Public Property Get InIDLE() As Boolean
    InIDLE = m_bInIDLE
End Property
'
'---------------------------------------------------------------------------------------
' Procedure : TimerProc
' Purpose   :
'---------------------------------------------------------------------------------------
Private Sub TimerProc(ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long)
    Dim lTime As Long
    
    If m_bGetLastInputInfo Then
        Dim tLASTINPUTINFO As LASTINPUTINFO
        tLASTINPUTINFO.lSize = Len(tLASTINPUTINFO)
        If Not (GetLastInputInfo(tLASTINPUTINFO) = ERROR) Then
            lTime = (GetTickCount - tLASTINPUTINFO.lTime) \ 1000
        End If
    Else
        lTime = (GetTickCount - m_lTime) \ 1000
    End If
    
    If lTime > m_lIDLETime Then
        If Not m_bInIDLE Then
            m_bInIDLE = True
            'At this point you can fire an event
            '---------------------------------------------------------------------------------------
            MDIForm1.Caption = "Hotel Management System   ::::IDLE::::"
            MDIForm1.Picture1.Visible = True
            MDIForm1.Timer1.Enabled = True
            MDIForm1.Image1.Picture = LoadPicture(App.Path & "\SCR\1.JPG")
            PlayWave App.Path & "\123\ELECTRIC MUSIC.wav"
            '---------------------------------------------------------------------------------------
        End If
    Else
        If m_bInIDLE Then
            m_bInIDLE = False
            'At this point you can fire an event
            '---------------------------------------------------------------------------------------
            MDIForm1.Caption = "Hotel Management System   ::::Not IDLE::::"
            MDIForm1.Picture1.Visible = False
            MDIForm1.Timer1.Enabled = False
            PlayWave App.Path & "\123\ELECTRIC MUSIC.wav"
            '---------------------------------------------------------------------------------------
        End If
    End If
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : KeyboardProc
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function KeyboardProc(ByVal idHook As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Not idHook < 0 Then
        m_lTime = GetTickCount
    End If
    KeyboardProc = CallNextHookEx(m_lKeyHook, idHook, wParam, ByVal lParam)
End Function
'
'---------------------------------------------------------------------------------------
' Procedure : MouseProc
' Purpose   :
'---------------------------------------------------------------------------------------
Public Function MouseProc(ByVal idHook As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
    If Not idHook < 0 Then
        m_lTime = GetTickCount
    End If
    MouseProc = CallNextHookEx(m_lMouseHook, idHook, wParam, ByVal lParam)
End Function
'
'---------------------------------------------------------------------------------------
' Procedure : Init
' Purpose   :
'---------------------------------------------------------------------------------------
Public Sub Init(lhWnd As Long, Optional lCheckTime As Long = 500)
    If GetVersion < 3 Then
        m_bGetLastInputInfo = True
    Else
        m_lKeyHook = SetWindowsHookEx(WH_KEYBOARD, AddressOf KeyboardProc, App.hInstance, App.ThreadID)
        m_lMouseHook = SetWindowsHookEx(WH_MOUSE, AddressOf MouseProc, App.hInstance, App.ThreadID)
    End If
    m_lhWnd = lhWnd
    SetTimer lhWnd, 0, lCheckTime, AddressOf TimerProc
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : Terminate
' Purpose   :
'---------------------------------------------------------------------------------------
Public Sub Terminate()
    KillTimer m_lhWnd, 0
    UnhookWindowsHookEx m_lKeyHook
    UnhookWindowsHookEx m_lMouseHook
End Sub
'
'---------------------------------------------------------------------------------------
' Procedure : GetVersion
' Purpose   :
'---------------------------------------------------------------------------------------
Private Function GetVersion() As Long
    Dim tOSVERSIONINFO As OSVERSIONINFO
    GetVersionEx tOSVERSIONINFO
    GetVersion = tOSVERSIONINFO.dwMinorVersion
End Function

