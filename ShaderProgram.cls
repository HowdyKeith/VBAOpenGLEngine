Option Explicit

' ============================================================
' ShaderProgram.cls  WEEK 1 PERFORMANCE PASS
' WEEK 1 CHANGES:
'   - Uniform location cache added (Scripting.Dictionary)
'     glGetUniformLocation is a DRIVER ROUND-TRIP costing ~1-5µs each call.
'     At 16 uniforms per object x 100 objects = 1600 calls/frame = ~5ms wasted.
'     Cache converts repeat calls to a dictionary lookup (~0.1µs).
'   - SetUniform3f was a stub - now implemented (was always a no-op)
'   - modPerf.CountUniform() called on every upload for profiling
'   - ClearCache() added for shader reload/recompile scenarios
' ============================================================

Private m_ProgramID   As Long
Private m_LocCache    As Object   ' Scripting.Dictionary: name -> Long location

' =========================
' PROPERTIES
' =========================
Public Property Get id() As Long: id = m_ProgramID: End Property

' =========================
' LIFECYCLE
' =========================
Private Sub Class_Initialize()
    Set m_LocCache = CreateObject("Scripting.Dictionary")
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub

' =========================
' CREATE
' =========================
Public Function CreateFromSource(ByVal vSrc As String, ByVal fSrc As String) As Boolean
    m_ProgramID = modShaderCompiler.Compile(vSrc, fSrc)
    CreateFromSource = (m_ProgramID <> 0)
    If m_ProgramID <> 0 Then
        m_LocCache.RemoveAll   ' clear stale cache on recompile
        Debug.Print "[ShaderProgram] Created ID=" & m_ProgramID
    End If
End Function

Public Function CreateDefault() As Boolean
    Dim vSrc As String, fSrc As String
    vSrc = "#version 330 core" & vbLf & _
           "layout (location = 0) in vec3 aPos;" & vbLf & _
           "void main() { gl_Position = vec4(aPos, 1.0); }"
    fSrc = "#version 330 core" & vbLf & _
           "out vec4 FragColor;" & vbLf & _
           "void main() { FragColor = vec4(1.0, 0.5, 0.2, 1.0); }"
    CreateDefault = CreateFromSource(vSrc, fSrc)
End Function

' =========================
' BIND / UNBIND
' =========================
Public Sub Use()
    If m_ProgramID <> 0 Then GL.glUseProgram m_ProgramID
End Sub

Public Sub Unuse()
    GL.glUseProgram 0
End Sub

Public Sub Destroy()
    If m_ProgramID <> 0 Then
        GL.glDeleteProgram m_ProgramID
        m_ProgramID = 0
    End If
    If Not m_LocCache Is Nothing Then m_LocCache.RemoveAll
End Sub

' WEEK 1: Clear location cache (call if shader is reloaded)
Public Sub ClearCache()
    If Not m_LocCache Is Nothing Then m_LocCache.RemoveAll
End Sub

' =========================
' UNIFORM SETTERS
' =========================
Public Sub SetUniform1i(ByVal name As String, ByVal v0 As Long)
    Dim loc As Long: loc = GetLocation(name)
    If loc = -1 Then Exit Sub
    GL.glUniform1i loc, v0
    modPerf.CountUniform
End Sub

Public Sub SetUniform1f(ByVal name As String, ByVal v0 As Single)
    Dim loc As Long: loc = GetLocation(name)
    If loc = -1 Then Exit Sub
    GL.glUniform1f loc, v0
    modPerf.CountUniform
End Sub

Public Sub SetUniform3f(ByVal name As String, ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Dim loc As Long: loc = GetLocation(name)
    If loc = -1 Then Exit Sub
    GL.glUniform3f loc, x, y, z
    modPerf.CountUniform
End Sub

Public Sub SetUniformMatrix4fv(ByVal name As String, ByRef mat As GLMatrix)
    Dim loc As Long: loc = GetLocation(name)
    If loc = -1 Then Exit Sub
    GL.glUniformMatrix4fv loc, 1, GL.GL_FALSE, mat.GetPtr()
    modPerf.CountUniform
End Sub

Public Sub SetUniformMat4(ByVal name As String, ByRef mat As GLMath.Mat4)
    Dim loc As Long: loc = GetLocation(name)
    If loc = -1 Then Exit Sub
    GL.glUniformMatrix4fv loc, 1, GL.GL_FALSE, VarPtr(mat.m(0))
    modPerf.CountUniform
End Sub

' =========================
' WEEK 1: CACHED LOCATION LOOKUP
' The first call to each name does the driver round-trip.
' Every subsequent call is a dictionary lookup (~0.1µs vs ~3µs).
' =========================
Private Function GetLocation(ByVal name As String) As Long
    If m_ProgramID = 0 Then GetLocation = -1: Exit Function

    ' WEEK 1: Check cache first
    If m_LocCache.Exists(name) Then
        GetLocation = m_LocCache(name)
        Exit Function
    End If

    ' Not cached - do the driver round-trip once
    Dim b() As Byte
    b = StrConv(name & vbNullChar, vbFromUnicode)
    Dim loc As Long
    loc = GL.glGetUniformLocation(m_ProgramID, VarPtr(b(0)))

    ' Cache even -1 results (uniform doesn't exist / was optimised out)
    m_LocCache(name) = loc
    GetLocation = loc
End Function

' =========================
' DEBUG
' =========================
Public Sub DebugPrintCache()
    Debug.Print "[ShaderProgram " & m_ProgramID & "] Location cache (" & m_LocCache.Count & " entries):"
    Dim k As Variant
    For Each k In m_LocCache.Keys
        Debug.Print "  " & k & " -> " & m_LocCache(k)
    Next k
End Sub
