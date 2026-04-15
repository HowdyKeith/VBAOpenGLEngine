Option Explicit

' ============================================================
' Class       : ComputeParticleSystem
' Version     : v1.1 FIXED
' FIXES:
'   - GL.glGetUniformLocation(prog, "deltaTime") - glGetUniformLocation
'     takes (prog, namePtr As LongPtr), not a String. VBA would pass the
'     BSTR pointer of the unicode string, which is wrong for a C API.
'     Fixed to use GL.glGetUniformLocationS(prog, "deltaTime") which is
'     the new string-safe overload added to GL.bas v9.4.
'   - GL_SHADER_STORAGE_BUFFER, GL_DYNAMIC_DRAW, GL_SHADER_STORAGE_BARRIER_BIT,
'     GL_POINTS were unqualified - now GL.GL_*
' ============================================================

Private m_ComputeProgram As Long
Private m_ParticleSSBO   As Long
Private m_ParticleCount  As Long

' ============================================================
' INITIALIZE
' ============================================================
Public Sub Initialize(ByVal maxParticles As Long)
    m_ParticleCount = maxParticles

    ' Load compute shader from embedded source
    ' (requires GL context to be current)
    m_ComputeProgram = CompileComputeShader()

    ' Allocate particle SSBO
    ' Each particle: vec4 position + vec4 velocity + float life + padding = 32 bytes
    GL.glGenBuffers 1, m_ParticleSSBO
    GL.glBindBuffer GL.GL_SHADER_STORAGE_BUFFER, m_ParticleSSBO
    GL.glBufferData GL.GL_SHADER_STORAGE_BUFFER, maxParticles * 32, 0, GL.GL_DYNAMIC_DRAW
    GL.glBindBufferBase GL.GL_SHADER_STORAGE_BUFFER, 0, m_ParticleSSBO
    GL.glBindBuffer GL.GL_SHADER_STORAGE_BUFFER, 0

    Debug.Print "[ComputeParticles] Initialized " & maxParticles & " particles."
End Sub

' ============================================================
' UPDATE (GPU compute pass)
' ============================================================
Public Sub Update(ByVal deltaTime As Single)
    If m_ComputeProgram = 0 Then Exit Sub

    GL.glUseProgram m_ComputeProgram

    ' FIXED: use glGetUniformLocationS (string overload) instead of
    ' glGetUniformLocation which takes a raw LongPtr (C char*)
    Dim dtLoc As Long
    dtLoc = GL.glGetUniformLocationS(m_ComputeProgram, "deltaTime")
    If dtLoc <> -1 Then GL.glUniform1f dtLoc, deltaTime

    ' Dispatch: 1 workgroup per 256 particles
    Dim groups As Long
    groups = (m_ParticleCount + 255) \ 256
    GL.glDispatchCompute groups, 1, 1

    ' FIXED: GL_SHADER_STORAGE_BARRIER_BIT now qualified as GL.GL_SHADER_STORAGE_BARRIER_BIT
    GL.glMemoryBarrier GL.GL_SHADER_STORAGE_BARRIER_BIT
End Sub

' ============================================================
' RENDER (simple point sprites)
' ============================================================
Public Sub Render()
    ' FIXED: GL_POINTS now GL.GL_POINTS
    GL.glBindBuffer GL.GL_SHADER_STORAGE_BUFFER, m_ParticleSSBO
    GL.glDrawArrays GL.GL_POINTS, 0, m_ParticleCount
    GL.glBindBuffer GL.GL_SHADER_STORAGE_BUFFER, 0
End Sub

' ============================================================
' CLEANUP
' ============================================================
Public Sub Destroy()
    If m_ParticleSSBO   <> 0 Then GL.glDeleteBuffers 1, m_ParticleSSBO:  m_ParticleSSBO = 0
    If m_ComputeProgram <> 0 Then GL.glDeleteProgram m_ComputeProgram:    m_ComputeProgram = 0
End Sub

' ============================================================
' PRIVATE: compile compute shader from EmbeddedShaders constant
' ============================================================
Private Function CompileComputeShader() As Long
    If GL.p_glCreateShader = 0 Then GL.LoadShaderExtensions

    Dim shader As Long
    shader = GL.glCreateShader(GL.GL_COMPUTE_SHADER)
    If shader = 0 Then
        Debug.Print "[ComputeParticles] glCreateShader(COMPUTE) returned 0"
        Exit Function
    End If

    GL.glShaderSource shader, EmbeddedShaders.PARTICLE_COMPUTE
    GL.glCompileShader shader

    Dim status As Long
    GL.glGetShaderiv shader, GL.GL_COMPILE_STATUS, status
    If status = 0 Then
        Debug.Print "[ComputeParticles] Compute shader compile failed."
        GL.glDeleteShader shader
        Exit Function
    End If

    Dim prog As Long
    prog = GL.glCreateProgram()
    GL.glAttachShader prog, shader
    GL.glLinkProgram prog
    GL.glDeleteShader shader

    Debug.Print "[ComputeParticles] Compute program ID=" & prog
    CompileComputeShader = prog
End Function
