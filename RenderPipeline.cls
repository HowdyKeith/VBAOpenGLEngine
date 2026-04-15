Option Explicit

' ============================================================
' Class       : RenderPipeline
' Version     : v9.4 FIXED
' FIXES:
'   - GL.glGenQueries, GL.glBeginQuery, GL.glEndQuery,
'     GL.glGetQueryObjectiv, GL.glDeleteQueries were called but
'     those functions didn't exist in old GL.bas - now added in v9.4
'   - GL_TIME_ELAPSED and GL_QUERY_RESULT were undefined constants - now in GL.bas
'   - Unqualified GL_COLOR_BUFFER_BIT / GL_DEPTH_BUFFER_BIT - now GL.GL_*
' ============================================================

Private m_Batcher          As MeshBatcher
Private m_Particles        As ComputeParticleSystem
Private m_TimerQuery       As Long
Private m_LastFrameTimeMS  As Double

' ============================================================
' INITIALIZE
' ============================================================
Public Sub Initialize()
    ' GL extension pointers must already be loaded (call after GL_Init)
    GL.GL_Init

    Set m_Batcher = New MeshBatcher
    m_Batcher.UseIndirectDrawing = True

    Set m_Particles = New ComputeParticleSystem
    m_Particles.Initialize 65536

    ' GPU Timer Query
    GL.glGenQueries 1, m_TimerQuery

    Debug.Print "[RenderPipeline] v9.4 Initialized - Indirect + Compute + Timer Ready"
End Sub

' ============================================================
' RENDER FRAME
' ============================================================
Public Sub RenderFrame(ByVal deltaTime As Single)
    ' Begin GPU timer
    If m_TimerQuery <> 0 Then
        GL.glBeginQuery GL.GL_TIME_ELAPSED, m_TimerQuery
    End If

    ' Clear
    GL.glClear GL.GL_COLOR_BUFFER_BIT Or GL.GL_DEPTH_BUFFER_BIT
    GL.glClearColor 0.1, 0.1, 0.15, 1#

    ' Update compute particle simulation
    If Not m_Particles Is Nothing Then
        m_Particles.Update deltaTime
    End If

    ' Flush geometry
    If Not m_Batcher Is Nothing Then
        If m_Batcher.UseIndirectDrawing Then
            m_Batcher.FlushIndirect
        Else
            m_Batcher.FlushDynamic
            m_Batcher.FlushInstanced
        End If
    End If

    ' Draw particles
    If Not m_Particles Is Nothing Then
        m_Particles.Render
    End If

    ' End GPU timer and read result
    If m_TimerQuery <> 0 Then
        GL.glEndQuery GL.GL_TIME_ELAPSED
        m_LastFrameTimeMS = GetLastFrameTimeMS()
    End If
End Sub

' ============================================================
' FRAME TIME QUERY
' ============================================================
Private Function GetLastFrameTimeMS() As Double
    If m_TimerQuery = 0 Then Exit Function
    Dim result As Long
    GL.glGetQueryObjectiv m_TimerQuery, GL.GL_QUERY_RESULT, result
    GetLastFrameTimeMS = result / 1000000#   ' nanoseconds -> milliseconds
End Function

Public Property Get LastFrameTimeMS() As Double
    LastFrameTimeMS = m_LastFrameTimeMS
End Property

' ============================================================
' SUBMIT GEOMETRY
' ============================================================
Public Sub SubmitMesh(ByVal vao As Long, ByVal indexCount As Long, _
                      Optional ByVal instanceCount As Long = 1)
    If Not m_Batcher Is Nothing Then
        m_Batcher.SubmitIndirect indexCount, instanceCount, 0, 0
    End If
End Sub

' ============================================================
' SHUTDOWN
' ============================================================
Public Sub Shutdown()
    If Not m_Batcher Is Nothing Then
        m_Batcher.Destroy
        Set m_Batcher = Nothing
    End If
    If Not m_Particles Is Nothing Then
        m_Particles.Destroy
        Set m_Particles = Nothing
    End If
    If m_TimerQuery <> 0 Then
        GL.glDeleteQueries 1, m_TimerQuery
        m_TimerQuery = 0
    End If
    Debug.Print "[RenderPipeline] Shutdown complete."
End Sub

Private Sub Class_Terminate()
    Shutdown
End Sub
