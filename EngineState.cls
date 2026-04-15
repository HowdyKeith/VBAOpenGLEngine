Option Explicit

' ============================================================
' EngineState.cls  WEEK 1 PERFORMANCE PASS
' WEEK 1 CHANGES:
'   - Timer() replaced with Win32GL.GetTime() (QPC-backed, microsecond precision)
'     Timer() has ~15.6ms resolution on Windows - causes uneven deltaTime and
'     can make physics/movement stutter at high frame rates
'   - deltaTime capped at 100ms to prevent spiral-of-death on stall frames
'   - TotalTime accumulated from QPC deltas (consistent with modPerf.TotalTime)
' ============================================================

Public Entities   As Object
Public Components As Object
Public Systems    As Object

Public physics    As PhysicsSystem
Public renderer   As Object
Public Camera     As Object
Public InputSys   As Object

Public deltaTime  As Single
Public TotalTime  As Double
Public IsRunning  As Boolean
Public IsPaused   As Boolean
Public IsDebug    As Boolean

' WEEK 1: QPC-backed timing state
Private m_LastTime As Double
Private m_QPCReady As Boolean

Private Sub Class_Initialize()
    Set Entities   = Nothing
    Set Components = Nothing
    Set Systems    = Nothing
    Set physics    = New PhysicsSystem
    Set renderer   = Nothing
    Set Camera     = Nothing
    Set InputSys   = Nothing

    IsRunning = False
    IsPaused  = False
    IsDebug   = False
    deltaTime = 0!
    TotalTime = 0#

    ' WEEK 1: Use QPC for initial timestamp
    m_LastTime = Win32GL.GetTime()
    m_QPCReady = True
End Sub

Public Sub StartEngine()
    IsRunning  = True
    ' WEEK 1: Reset QPC baseline on engine start
    m_LastTime = Win32GL.GetTime()
End Sub

Public Sub Update()
    If Not IsRunning Then Exit Sub

    ' WEEK 1: QPC delta time (was: Timer)
    Dim nowT As Double
    nowT = Win32GL.GetTime()

    deltaTime = CSng(nowT - m_LastTime)
    m_LastTime = nowT

    ' WEEK 1: Cap to 100ms - prevents physics explosion on stall frames
    If deltaTime > 0.1 Then deltaTime = 0.1!

    TotalTime = TotalTime + deltaTime

    If IsPaused Then Exit Sub

    If Not InputSys Is Nothing Then
        On Error Resume Next
        CallByName InputSys, "Update", VbMethod
        On Error GoTo 0
    End If

    If Not Camera Is Nothing Then
        On Error Resume Next
        CallByName Camera, "Update", VbMethod, deltaTime
        On Error GoTo 0
    End If

    If Not physics Is Nothing Then
        physics.Step deltaTime
    End If

    If Not Systems Is Nothing Then
        On Error Resume Next
        CallByName Systems, "Update", VbMethod, deltaTime
        On Error GoTo 0
    End If

    If Not renderer Is Nothing Then
        On Error Resume Next
        CallByName renderer, "Render", VbMethod, Camera
        On Error GoTo 0
    End If
End Sub

Public Sub StopEngine():   IsRunning = False:  End Sub
Public Sub PauseEngine():  IsPaused  = True:   End Sub
Public Sub ToggleDebug():  IsDebug   = Not IsDebug: End Sub

Public Sub ResumeEngine()
    IsPaused   = False
    m_LastTime = Win32GL.GetTime()   ' WEEK 1: reset baseline on resume
End Sub

Public Sub BindPhysics(ByRef phys As PhysicsSystem):  Set physics  = phys:   End Sub
Public Sub BindRenderer(ByRef rend As Object):         Set renderer = rend:   End Sub
Public Sub BindCamera(ByRef cam As Object):            Set Camera   = cam:    End Sub
Public Sub BindInput(ByRef InputSystem As Object):     Set InputSys = InputSystem: End Sub

Public Sub Reset()
    Set physics  = New PhysicsSystem
    deltaTime    = 0!
    TotalTime    = 0#
    m_LastTime   = Win32GL.GetTime()
End Sub

Public Sub DebugPrint()
    Debug.Print "---- ENGINE STATE ----"
    Debug.Print "Running: ";   IsRunning
    Debug.Print "Paused: ";    IsPaused
    Debug.Print "DeltaTime: "; deltaTime
    Debug.Print "TotalTime: "; TotalTime
    If Not physics Is Nothing Then Debug.Print "Gravity: "; physics.Gravity
End Sub
