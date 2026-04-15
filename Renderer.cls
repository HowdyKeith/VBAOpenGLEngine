Option Explicit

' ============================================================
' Renderer.cls  WEEK 1 PERFORMANCE PASS
' WEEK 1 CHANGES:
'   - State tracking extended to cover blend mode, depth test, and cull state.
'     Every GL state call was unconditional before - now each checks a dirty flag.
'     A scene with N objects that all use the same material saved N-1 driver calls
'     per state category per frame.
'   - modPerf.CountStateChange() and CountDraw() integrated for profiling
'   - SetMaterial() added: one call sets all state for a material in one place,
'     minimising redundant comparisons when rendering many objects
'   - Viewport dirty flag: SetViewport only calls driver if dimensions changed
' ============================================================

' --- Current bound state ---
Private m_CurrentShader  As Long
Private m_CurrentVAO     As Long
Private m_CurrentTexture As Long

' --- WEEK 1: Extended state tracking ---
Private m_DepthEnabled   As Long   ' -1 = unknown, 0 = off, 1 = on
Private m_BlendEnabled   As Long
Private m_CullEnabled    As Long
Private m_CullMode       As Long
Private m_BlendSrc       As Long
Private m_BlendDst       As Long
Private m_WireframeOn    As Long

Private m_VpX As Long, m_VpY As Long, m_VpW As Long, m_VpH As Long

Private Sub Class_Initialize()
    ResetState
End Sub

' ============================================================
' CLEAR
' ============================================================
Public Sub Clear(ByVal r As Single, ByVal g As Single, ByVal b As Single)
    modGL.glClearColor r, g, b, 1#
    modGL.glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
End Sub

' ============================================================
' SHADER  (state-change guarded)
' ============================================================
Public Sub UseShader(ByVal programID As Long)
    If m_CurrentShader = programID Then Exit Sub
    modGL.glUseProgram programID
    m_CurrentShader = programID
    modPerf.CountStateChange
End Sub

' ============================================================
' VAO  (state-change guarded)
' ============================================================
Public Sub BindVAO(ByVal vao As Long)
    If m_CurrentVAO = vao Then Exit Sub
    modGL.glBindVertexArray vao
    m_CurrentVAO = vao
    modPerf.CountStateChange
End Sub

' ============================================================
' TEXTURE  (state-change guarded)
' ============================================================
Public Sub BindTexture(ByVal textureID As Long, Optional ByVal slot As Long = 0)
    If m_CurrentTexture = textureID Then Exit Sub
    modGL.glActiveTexture GL_TEXTURE0 + slot
    modGL.glBindTexture GL_TEXTURE_2D, textureID
    m_CurrentTexture = textureID
    modPerf.CountStateChange
End Sub

' ============================================================
' DRAW CALLS
' ============================================================
Public Sub DrawIndexed(ByVal indexCount As Long)
    If indexCount <= 0 Then Exit Sub
    modGL.glDrawElements GL_TRIANGLES, indexCount, GL_UNSIGNED_INT, 0
    modPerf.CountDraw indexCount \ 3
End Sub

Public Sub DrawArrays(ByVal vertexCount As Long)
    If vertexCount <= 0 Then Exit Sub
    modGL.glDrawArrays GL_TRIANGLES, 0, vertexCount
    modPerf.CountDraw vertexCount \ 3
End Sub

Public Sub DrawInstanced(ByVal indexCount As Long, ByVal instanceCount As Long)
    If indexCount <= 0 Or instanceCount <= 0 Then Exit Sub
    modGL.glDrawElementsInstanced GL_TRIANGLES, indexCount, GL_UNSIGNED_INT, 0, instanceCount
    modPerf.CountDraw (indexCount \ 3) * instanceCount
End Sub

' ============================================================
' WEEK 1: DEPTH STATE (guarded)
' ============================================================
Public Sub EnableDepth()
    If m_DepthEnabled = 1 Then Exit Sub
    modGL.glEnable GL_DEPTH_TEST
    m_DepthEnabled = 1
    modPerf.CountStateChange
End Sub

Public Sub DisableDepth()
    If m_DepthEnabled = 0 Then Exit Sub
    modGL.glDisable GL_DEPTH_TEST
    m_DepthEnabled = 0
    modPerf.CountStateChange
End Sub

' ============================================================
' WEEK 1: BLEND STATE (guarded)
' ============================================================
Public Sub EnableBlend(Optional ByVal src As Long = GL_SRC_ALPHA, _
                       Optional ByVal dst As Long = GL_ONE_MINUS_SRC_ALPHA)
    If m_BlendEnabled = 1 And m_BlendSrc = src And m_BlendDst = dst Then Exit Sub
    modGL.glEnable GL_BLEND
    modGL.glBlendFunc src, dst
    m_BlendEnabled = 1: m_BlendSrc = src: m_BlendDst = dst
    modPerf.CountStateChange
End Sub

Public Sub DisableBlend()
    If m_BlendEnabled = 0 Then Exit Sub
    modGL.glDisable GL_BLEND
    m_BlendEnabled = 0
    modPerf.CountStateChange
End Sub

' ============================================================
' WEEK 1: CULL STATE (guarded)
' ============================================================
Public Sub EnableCull()
    If m_CullEnabled = 1 Then Exit Sub
    modGL.glEnable GL_CULL_FACE
    m_CullEnabled = 1
    modPerf.CountStateChange
End Sub

Public Sub DisableCull()
    If m_CullEnabled = 0 Then Exit Sub
    modGL.glDisable GL_CULL_FACE
    m_CullEnabled = 0
    modPerf.CountStateChange
End Sub

Public Sub SetCullMode(ByVal mode As Long)
    If m_CullMode = mode Then Exit Sub
    modGL.glCullFace mode
    m_CullMode = mode
    modPerf.CountStateChange
End Sub

' ============================================================
' WEEK 1: WIREFRAME (guarded)
' ============================================================
Public Sub SetWireframe(ByVal enabled As Boolean)
    Dim want As Long: want = IIf(enabled, 1, 0)
    If m_WireframeOn = want Then Exit Sub
    If enabled Then
        modGL.glPolygonMode GL_FRONT_AND_BACK, GL_LINE
    Else
        modGL.glPolygonMode GL_FRONT_AND_BACK, GL_FILL
    End If
    m_WireframeOn = want
    modPerf.CountStateChange
End Sub

' ============================================================
' WEEK 1: SetMaterial - apply all material state in one guarded call
' Prevents N redundant state comparisons when N objects share a material
' ============================================================
Public Sub SetMaterial(ByVal shaderID As Long, ByVal texID As Long, _
                       ByVal blendMode As Long, ByVal depthTest As Boolean, _
                       ByVal wireframe As Boolean)
    UseShader shaderID
    If texID >= 0 Then BindTexture texID

    If depthTest Then EnableDepth Else DisableDepth

    Select Case blendMode
        Case 0: DisableBlend
        Case 1: EnableBlend GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA
        Case 2: EnableBlend GL_SRC_ALPHA, GL_ONE
    End Select

    SetWireframe wireframe
End Sub

' ============================================================
' WEEK 1: VIEWPORT (guarded - only calls driver when dimensions change)
' ============================================================
Public Sub SetViewport(ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long)
    If m_VpX = x And m_VpY = y And m_VpW = w And m_VpH = h Then Exit Sub
    modGL.glViewport x, y, w, h
    m_VpX = x: m_VpY = y: m_VpW = w: m_VpH = h
    modPerf.CountStateChange
End Sub

' ============================================================
' ERROR CHECK
' ============================================================
Public Sub CheckError(ByVal tag As String)
    Dim errCode As Long: errCode = modGL.glGetError()
    If errCode <> 0 Then Debug.Print "[GL ERROR] " & tag & " Code: " & errCode
End Sub

' ============================================================
' RESET (marks all state as unknown, forces re-upload next draw)
' ============================================================
Public Sub ResetState()
    m_CurrentShader  = 0
    m_CurrentVAO     = 0
    m_CurrentTexture = 0
    m_DepthEnabled   = -1   ' -1 = unknown
    m_BlendEnabled   = -1
    m_CullEnabled    = -1
    m_CullMode       = -1
    m_BlendSrc       = -1
    m_BlendDst       = -1
    m_WireframeOn    = -1
    m_VpX = -1: m_VpY = -1: m_VpW = -1: m_VpH = -1
    modGL.glBindVertexArray 0
    modGL.glUseProgram 0
End Sub
