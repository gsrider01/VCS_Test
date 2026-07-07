Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Compare Database

Private m_Elapsed As Double
Private m_StartTick As Double
Private m_Running As Boolean

Private Sub Form_Load()
    Me.TimerInterval = 0
    ResetStopwatch
End Sub

Private Sub cmdStart_Click()
    If Not m_Running Then
        m_StartTick = Timer
        m_Running = True
        Me.TimerInterval = 50
        SetButtonStates
    End If
End Sub

Private Sub cmdStop_Click()
    If m_Running Then
        m_Elapsed = m_Elapsed + ElapsedSinceStart()
        m_Running = False
        Me.TimerInterval = 0
        UpdateDisplay
        SetButtonStates
    End If
End Sub

Private Sub cmdReset_Click()
    m_Elapsed = 0
    m_Running = False
    Me.TimerInterval = 0
    UpdateDisplay
    SetButtonStates
End Sub

Private Sub Form_Timer()
    UpdateDisplay
End Sub

Private Sub ResetStopwatch()
    m_Elapsed = 0
    m_Running = False
    Me.TimerInterval = 0
    Me.txtTime = "00:00.00"
    SetButtonStates
End Sub

Private Sub UpdateDisplay()
    Dim total As Double

    If m_Running Then
        total = m_Elapsed + ElapsedSinceStart()
    Else
        total = m_Elapsed
    End If

    Me.txtTime = FormatElapsed(total)
End Sub

Private Function ElapsedSinceStart() As Double
    Dim delta As Double

    delta = Timer - m_StartTick
    If delta < 0 Then delta = delta + 86400#
    ElapsedSinceStart = delta
End Function

Private Function FormatElapsed(ByVal seconds As Double) As String
    Dim mins As Long
    Dim secs As Long
    Dim cents As Long
    Dim wholeSecs As Long

    wholeSecs = Int(seconds)
    mins = wholeSecs \ 60
    secs = wholeSecs Mod 60
    cents = Int((seconds - wholeSecs) * 100)
    If cents > 99 Then cents = 99

    FormatElapsed = Format$(mins, "00") & ":" & Format$(secs, "00") & "." & Format$(cents, "00")
End Function

Private Sub SetButtonStates()
    Me.cmdStart.Enabled = Not m_Running
    Me.cmdStop.Enabled = m_Running
    Me.cmdReset.Enabled = True
End Sub
