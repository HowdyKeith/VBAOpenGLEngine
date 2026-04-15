Option Explicit

' ============================================================
' modPerf.bas - v1.0  WEEK 1: PERFORMANCE MONITOR
' Central performance counter for FPS, frame time, draw calls,
' state changes, and uniform uploads. Zero overhead when disabled.
' ============================================================

' --- Config ---
Public Const PERF_ENABLED       As Boolean = True   ' set False to zero all overhead
Public Const PERF_SMOOTH_FRAMES As Long = 60         ' rolling average window

' --- Frame timing ---
Private m_FreqInit  As Boolean
Private m_Freq      As Currency
Private m_T0        As Currency        ' start of current frame
Private m_T1        As Currency        ' end of previous frame

Private m_FrameTime As Double          ' last frame seconds
Private m_Smooth()  As Double          ' ring buffer of frame times
Private m_SmoothIdx As Long

' --- Counters (reset each frame) ---
Public DrawCalls    As Long
Public StateChanges As Long
Public UniformUploads As Long
Public VerticesDrawn As Long
Public TrianglesDrawn As Long

' --- Accumulated ---
Public TotalFrames  As Long
Public TotalTime    As Double

' --- Output ---
Public FPS          As Single
Public FrameTimeMS  As Single
Public FrameTimeAvgMS As Single

' ============================================================
' INIT
' ============================================================
Public Sub Init()
    If Not PERF_ENABLED Then Exit Sub
    ReDim m_Smooth(0 To PERF_SMOOTH_FRAMES - 1)
    m_SmoothIdx = 0
    If Not m_FreqInit Then
        Win32GL.QueryPerformanceFrequency m_Freq
        m_FreqInit = True
    End If
    Win32GL.QueryPerformanceCounter m_T1
    Debug.Print "[Perf] Monitor initialised. Smooth window=" & PERF_SMOOTH_FRAMES & " frames."
End Sub

' ============================================================
' CALL AT START OF EACH FRAME
' ============================================================
Public Sub BeginFrame()
    If Not PERF_ENABLED Then Exit Sub
    Win32GL.QueryPerformanceCounter m_T0
    ' Reset per-frame counters
    DrawCalls     = 0
    StateChanges  = 0
    UniformUploads = 0
    VerticesDrawn = 0
    TrianglesDrawn = 0
End Sub

' ============================================================
' CALL AT END OF EACH FRAME (after SwapBuffers)
' ============================================================
Public Sub EndFrame()
    If Not PERF_ENABLED Then Exit Sub

    Dim t As Currency
    Win32GL.QueryPerformanceCounter t
    m_FrameTime = CDbl(t - m_T0) / CDbl(m_Freq)

    ' Rolling average
    m_Smooth(m_SmoothIdx) = m_FrameTime
    m_SmoothIdx = (m_SmoothIdx + 1) Mod PERF_SMOOTH_FRAMES

    Dim total As Double
    Dim i As Long
    For i = 0 To PERF_SMOOTH_FRAMES - 1: total = total + m_Smooth(i): Next i
    Dim avg As Double: avg = total / PERF_SMOOTH_FRAMES

    ' Published values
    FrameTimeMS    = CSng(m_FrameTime * 1000#)
    FrameTimeAvgMS = CSng(avg * 1000#)
    FPS            = CSng(IIf(avg > 0, 1# / avg, 0))

    TotalFrames = TotalFrames + 1
    TotalTime   = TotalTime + m_FrameTime
End Sub

' ============================================================
' COUNTER HELPERS (called from Renderer / ShaderProgram etc.)
' ============================================================
Public Sub CountDraw(ByVal triangles As Long)
    If Not PERF_ENABLED Then Exit Sub
    DrawCalls     = DrawCalls + 1
    TrianglesDrawn = TrianglesDrawn + triangles
End Sub

Public Sub CountStateChange()
    If Not PERF_ENABLED Then StateChanges = StateChanges + 1
End Sub

Public Sub CountUniform()
    If Not PERF_ENABLED Then Exit Sub
    UniformUploads = UniformUploads + 1
End Sub

' ============================================================
' UPDATE WINDOW TITLE (call once per second or so)
' ============================================================
Private m_TitleTimer As Double

Public Sub UpdateWindowTitle(ByVal hWnd As LongPtr)
    If Not PERF_ENABLED Then Exit Sub
    TotalTime = TotalTime   ' read already set in EndFrame
    ' Rate-limit to once per second
    Static lastUpdate As Double
    If TotalTime - lastUpdate < 1# Then Exit Sub
    lastUpdate = TotalTime

    Dim title As String
    title = "VBA OpenGL Engine  |  " & _
            Format(FPS, "0.0") & " FPS  |  " & _
            Format(FrameTimeAvgMS, "0.00") & " ms  |  " & _
            "Draws: " & DrawCalls & "  Tris: " & TrianglesDrawn

    Win32GL.SetWindowText hWnd, title
End Sub

' ============================================================
' DEBUG DUMP
' ============================================================
Public Sub DebugPrint()
    If Not PERF_ENABLED Then Debug.Print "[Perf] Disabled": Exit Sub
    Debug.Print "---- PERF ----"
    Debug.Print "  FPS:          " & Format(FPS, "0.0")
    Debug.Print "  Frame (last): " & Format(FrameTimeMS, "0.00") & " ms"
    Debug.Print "  Frame (avg):  " & Format(FrameTimeAvgMS, "0.00") & " ms"
    Debug.Print "  Draw calls:   " & DrawCalls
    Debug.Print "  Triangles:    " & TrianglesDrawn
    Debug.Print "  State chg:    " & StateChanges
    Debug.Print "  Uniform ups:  " & UniformUploads
    Debug.Print "  Total frames: " & TotalFrames
End Sub
