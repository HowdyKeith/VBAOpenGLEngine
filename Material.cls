Option Explicit

' ============================================================
' Class       : Material
' Version     : v8.17 FIXED
' FIXES:
'   - GL_DEPTH_TEST, GL_BLEND, GL_FRONT_AND_BACK, GL_LINE, GL_FILL,
'     GL_SRC_ALPHA, GL_ONE_MINUS_SRC_ALPHA, GL_ONE were unqualified -
'     caused "Variable not defined" or ambiguous name errors.
'     All now use GL.GL_* explicit qualification.
'   - Apply() had "If shader Is Nothing Then Exit Sub" which would
'     skip shader bind for all materials that don't pass a shader object.
'     shaderID bind now happens unconditionally (before the Nothing check).
' ============================================================

' --- IDENTITY ---
Public id     As Long
Public name   As String
Public Active As Boolean

' --- CORE REFERENCES ---
Public MaterialSystemRef As Object

' --- SHADER ---
Public shaderID As Long

' --- TEXTURES ---
Public DiffuseTextureID  As Long
Public NormalTextureID   As Long
Public SpecularTextureID As Long

' --- SURFACE PROPERTIES ---
Public Shininess  As Single
Public Opacity    As Single
Public Metallic   As Single
Public Roughness  As Single

' --- RENDER FLAGS ---
Public UseLighting     As Boolean
Public UseTextures     As Boolean
Public Wireframe       As Boolean
Public DepthTest       As Boolean
Public CastShadows     As Boolean
Public ReceiveShadows  As Boolean

' --- BLENDING ---
' 0 = Opaque, 1 = AlphaBlend, 2 = Additive
Public BlendMode As Long

' ============================================================
' INIT
' ============================================================
Private Sub Class_Initialize()
    id     = -1
    name   = "New Material"
    Active = True

    shaderID          = -1
    DiffuseTextureID  = -1
    NormalTextureID   = -1
    SpecularTextureID = -1

    Shininess = 32#
    Opacity   = 1#
    Metallic  = 0#
    Roughness = 0.5

    UseLighting    = True
    UseTextures    = True
    Wireframe      = False
    DepthTest      = True
    CastShadows    = True
    ReceiveShadows = True
    BlendMode      = 0
End Sub

' ============================================================
' APPLY MATERIAL  (FIXED: all constants qualified)
' ============================================================
Public Sub Apply(ByRef shader As Object)
    If Not Active Then Exit Sub

    ' 1. SHADER BIND (always - even if shader object is Nothing we still use shaderID)
    If shaderID >= 0 Then GL.glUseProgram shaderID

    ' 2. DEPTH TEST
    If DepthTest Then
        GL.glEnable GL.GL_DEPTH_TEST
    Else
        GL.glDisable GL.GL_DEPTH_TEST
    End If

    ' 3. WIREFRAME
    If Wireframe Then
        GL.glPolygonMode GL.GL_FRONT_AND_BACK, GL.GL_LINE
    Else
        GL.glPolygonMode GL.GL_FRONT_AND_BACK, GL.GL_FILL
    End If

    ' 4. BLENDING
    Select Case BlendMode
        Case 0  ' Opaque
            GL.glDisable GL.GL_BLEND
        Case 1  ' Alpha Blend
            GL.glEnable GL.GL_BLEND
            GL.glBlendFunc GL.GL_SRC_ALPHA, GL.GL_ONE_MINUS_SRC_ALPHA
        Case 2  ' Additive
            GL.glEnable GL.GL_BLEND
            GL.glBlendFunc GL.GL_SRC_ALPHA, GL.GL_ONE
    End Select

    ' 5. UNIFORMS (only if a ShaderProgram object was passed)
    If shader Is Nothing Then Exit Sub

    If UseTextures Then
        If DiffuseTextureID  <> -1 Then shader.SetUniform1i "uDiffuseMap",  0
        If NormalTextureID   <> -1 Then shader.SetUniform1i "uNormalMap",   1
        If SpecularTextureID <> -1 Then shader.SetUniform1i "uSpecularMap", 2
    End If

    shader.SetUniform1f "uShininess",   Shininess
    shader.SetUniform1f "uOpacity",     Opacity
    shader.SetUniform1f "uMetallic",    Metallic
    shader.SetUniform1f "uRoughness",   Roughness
    shader.SetUniform1i "uUseLighting", IIf(UseLighting, 1, 0)
End Sub

' ============================================================
' HELPERS
' ============================================================
Public Sub SetOpaque()
    BlendMode = 0
    Opacity   = 1#
End Sub

Public Sub SetTransparent(ByVal alpha As Single)
    BlendMode = 1
    Opacity   = alpha
End Sub

Public Sub SetAdditive()
    BlendMode = 2
End Sub

' ============================================================
' DEBUG
' ============================================================
Public Sub DebugPrint()
    Debug.Print "---- [MATERIAL: " & name & "] ----"
    Debug.Print "  ID: " & id & " | ShaderID: " & shaderID
    Debug.Print "  Wireframe: " & Wireframe & " | Blend: " & BlendMode
    Debug.Print "  Opacity: " & Opacity & " | Roughness: " & Roughness
End Sub
