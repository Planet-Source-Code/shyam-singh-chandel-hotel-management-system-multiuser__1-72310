Attribute VB_Name = "WavePlay"
'Option Explicit

    Private Declare Function mciSendCommandA Lib "WinMM" _
        (ByVal wDeviceID As Long, ByVal Message As Long, _
        ByVal dwParam1 As Long, dwParam2 As Any) As Long

    Const MCI_OPEN = &H803
    Const MCI_CLOSE = &H804
    Const MCI_PLAY = &H806
    Const MCI_OPEN_TYPE = &H2000&
    Const MCI_OPEN_ELEMENT = &H200&
    Const MCI_WAIT = &H2&
    
    Private Type MCI_WAVE_OPEN_PARMS
        dwCallback As Long
        wDeviceID As Long
        lpstrDeviceType As String
        lpstrElementName As String
        lpstrAlias As String
        dwBufferSeconds As Long
    End Type
    
    Private Type MCI_PLAY_PARMS
        dwCallback As Long
        dwFrom As Long
        dwTo As Long
    End Type

Sub PlayWave(WaveFile As String)

    Dim errorCode As Integer
    Dim returnStr As Integer
    Dim errorStr As String * 256
    Dim MCIWaveOpenParms As MCI_WAVE_OPEN_PARMS
    Dim MCIPlayParms As MCI_PLAY_PARMS
    
    MCIWaveOpenParms.dwCallback = 0
    MCIWaveOpenParms.wDeviceID = 0
        
    MCIWaveOpenParms.lpstrDeviceType = "waveaudio"
    MCIWaveOpenParms.lpstrElementName = WaveFile
    
    MCIWaveOpenParms.lpstrAlias = 0
    MCIWaveOpenParms.dwBufferSeconds = 0
    
    errorCode = mciSendCommandA(0, MCI_OPEN, MCI_OPEN_TYPE Or MCI_OPEN_ELEMENT, _
                                MCIWaveOpenParms)
    
    If errorCode = 0 Then
        MCIPlayParms.dwCallback = 0
        MCIPlayParms.dwFrom = 0
        MCIPlayParms.dwTo = 0
    
        errorCode = mciSendCommandA(MCIWaveOpenParms.wDeviceID, MCI_PLAY, _
                                    MCI_WAIT, MCIPlayParms)
                                
        errorCode = mciSendCommandA(MCIWaveOpenParms.wDeviceID, MCI_CLOSE, _
                                    0, 0)
    End If
End Sub
        
'Private Sub Command1_Click()
'    PlayWave App.Path & "\123\ERROR LAFING.wav"
'End Sub

'Private Sub Sound_Click()
'    PlayWave App.Path & "\123\GOLF GROUND MUSIC.wav"
'End Sub

        


