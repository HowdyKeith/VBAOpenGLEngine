Option Explicit

' ============================================================
' modMain.bas  WEEK 1 PERFORMANCE PASS
' WEEK 1 CHANGES:
'   - DoEvents REMOVED from all loops (was killing frame rate by
'     yielding to Excel's message pump every frame ~15ms stall)
'   - Timer() REPLACED with QPC-backed Win32GL.GetTime() everywhere
'     (Timer has ~15ms resolution; QPC gives microsecond precision)
'   - modPerf integrated: BeginFrame/EndFrame wrap every loop tick
'   - Window title updated with live FPS/frametime once per second
'   - Frame cap added (optional, default 0 = uncapped)
'   - Escape key checked via Win32GL.IsKeyDown (no Excel dependency)
' ============================================================

#If Win64 Then
    Private Declare PtrSafe Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpLibFileName As String) As LongPtr
#Else
    Private Declare Function GetModuleHandle Lib "kernel32" Alias "GetModuleHandleA" (ByVal lpLibFileName As String) As Long
#End If

' Optional frame rate cap (0 = uncapped, 60 = 60fps cap, etc.)
Public Const TARGET_FPS As Long = 0

' ============================================================
' ENGINE DEMO - Spinning cube with FPS camera
' ============================================================
Public Sub RunEngineDemo()

    ' --- Window + Context ---
    Dim hWnd As LongPtr
    hWnd = modWindowManager.CreateGLWindow("VBA OpenGL Engine Demo", 800, 600)
    If hWnd = 0 Then MsgBox "Failed to create OpenGL window!", vbCritical: Exit Sub

    ' --- GL Pointers (must be after wglMakeCurrent) ---
    Dim hLib As LongPtr
    hLib = GetModuleHandle("opengl32.dll")
    If hLib = 0 Then
        MsgBox "opengl32.dll not found!", vbCritical
        modWindowManager.CloseGLWindow
        Exit Sub
    End If
    GLLoader.Init hLib
    GL.GL_Init

    ' --- Performance monitor ---
    modPerf.Init

    ' --- One-time GL state ---
    GL.glEnable GL.GL_DEPTH_TEST
    GL.glEnable GL.GL_CULL_FACE
    GL.glCullFace GL.GL_BACK

    ' --- Scene ---
    Dim cam As New Camera
    cam.SetPosition 0, 1.5, 5
    cam.FOV = 60: cam.AspectRatio = 800 / 600
    cam.NearPlane = 0.1: cam.FarPlane = 500

    Dim cube As New GLCube
    cube.SetPosition 0, 0, 0
    cube.SetScale 1, 1, 1

    Dim inputMgr As New InputManager
    Dim player As New FPSController
    player.BindInput inputMgr
    player.Position.SetValues 0, 1.5, 5

    ' --- WEEK 1: QPC timing (replaces Timer) ---
    Dim t0 As Double: t0 = Win32GL.GetTime()
    Dim totalTime As Double: totalTime = 0
    Dim dt As Single

    ' Frame cap setup
    Dim minFrameTime As Double
    minFrameTime = IIf(TARGET_FPS > 0, 1# / TARGET_FPS, 0)

    ' --- Main loop ---
    Dim running As Boolean: running = True
    Do While running

        ' WEEK 1: QPC delta time
        Dim t1 As Double: t1 = Win32GL.GetTime()
        dt = CSng(t1 - t0)
        If dt > 0.1 Then dt = 0.1   ' spike cap
        t0 = t1
        totalTime = totalTime + dt

        ' WEEK 1: Performance frame begin
        modPerf.BeginFrame

        ' Input
        inputMgr.Update
        If Win32GL.IsKeyDown(Win32GL.VK_ESCAPE) Then running = False

        ' Player + camera
        player.Update dt
        cam.SetPosition player.Position.x, player.Position.y, player.Position.z
        cam.Rotation.x = player.Camera.Rotation.x
        cam.Rotation.y = player.Camera.Rotation.y

        ' Spin cube
        cube.SetRotation CSng(totalTime * 30), CSng(totalTime * 45), 0

        ' --- Render ---
        GL.glClearColor 0.12, 0.14, 0.18, 1#
        GL.glClear GL.GL_COLOR_BUFFER_BIT Or GL.GL_DEPTH_BUFFER_BIT

        cam.ApplyProjection
        cam.ApplyView
        cube.Draw

        ' --- Present ---
        ' WEEK 1: PumpMessages only - NO DoEvents
        running = Win32GL.PumpMessages() And running
        modWindowManager.PageFlip

        ' WEEK 1: Performance frame end + window title update
        modPerf.EndFrame
        modPerf.UpdateWindowTitle hWnd

        ' WEEK 1: Optional frame cap (spin-wait for precise timing)
        If minFrameTime > 0 Then
            Do While Win32GL.GetTime() - t0 < minFrameTime: Loop
        End If

    Loop

    modWindowManager.CloseGLWindow
    Debug.Print "[EngineDemo] Exited. Total frames: " & modPerf.TotalFrames
    modPerf.DebugPrint
End Sub

' ============================================================
' FIRST PIXEL - Simplest triangle demo
' ============================================================
Public Sub RunFirstPixel()
    FirstPixel.RunFirstPixel
End Sub

' ============================================================
' MINIMAL START (bare clear loop)
' ============================================================
Public Sub StartEngine()
    Dim hWnd As LongPtr
    hWnd = modWindowManager.CreateGLWindow("VBA OpenGL Engine v2.0", 800, 600)
    If hWnd = 0 Then MsgBox "Failed to create OpenGL window!", vbCritical: Exit Sub

    Dim hLib As LongPtr
    hLib = GetModuleHandle("opengl32.dll")
    If hLib = 0 Then modWindowManager.CloseGLWindow: Exit Sub
    GLLoader.Init hLib
    GL.GL_Init
    modPerf.Init

    GL.glEnable GL.GL_DEPTH_TEST

    Dim running As Boolean: running = True
    Do While running
        modPerf.BeginFrame
        GL.glClearColor 0.2, 0.3, 0.3, 1#
        GL.glClear GL.GL_COLOR_BUFFER_BIT Or GL.GL_DEPTH_BUFFER_BIT
        ' --- render logic here ---
        running = Win32GL.PumpMessages()   ' NO DoEvents
        modWindowManager.PageFlip
        modPerf.EndFrame
        modPerf.UpdateWindowTitle hWnd
    Loop

    modWindowManager.CloseGLWindow
End Sub
