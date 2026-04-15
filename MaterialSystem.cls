Option Explicit

' ============================================================
' Class       : MaterialSystem
' Version     : v8.16
' Description : Render State + Shader + Texture + Lighting Binder
'               Synchronized with GL.bas v8.58
' ============================================================

' =========================
' MATERIAL STORAGE
' =========================

Private Type material
    name As String
    shaderID As Long

    DiffuseTextureID As Long
    NormalTextureID As Long
    SpecularTextureID As Long

    UseLighting As Boolean
    UseTextures As Boolean
    Wireframe As Boolean

    BlendMode As Long
    DepthTest As Boolean

    Shininess As Single
End Type

Private m_Materials() As material
Private m_Count As Long

' =========================
' LIGHT LINKS (GLOBAL)
' =========================

Public PointLightCount As Long
Public SpotLightCount As Long

' =========================
' INIT / TERMINATE
' =========================

Private Sub Class_Initialize()
    m_Count = 0
    ReDim m_Materials(0 To 0)

    PointLightCount = 0
    SpotLightCount = 0
End Sub

' =========================
' CREATE MATERIAL
' =========================

Public Function CreateMaterial(ByVal name As String, _
                               ByVal shaderID As Long) As Long
    Dim id As Long

    m_Count = m_Count + 1
    ReDim Preserve m_Materials(0 To m_Count - 1)

    id = m_Count - 1

    With m_Materials(id)
        .name = name
        .shaderID = shaderID

        .DiffuseTextureID = -1
        .NormalTextureID = -1
        .SpecularTextureID = -1

        .UseLighting = True
        .UseTextures = True
        .Wireframe = False

        .BlendMode = 0
        .DepthTest = True

        .Shininess = 32#
    End With

    CreateMaterial = id
End Function

' =========================
' SETTERS
' =========================

Public Sub SetDiffuseTexture(ByVal matID As Long, ByVal texID As Long)
    If Not Valid(matID) Then Exit Sub
    m_Materials(matID).DiffuseTextureID = texID
End Sub

Public Sub SetNormalTexture(ByVal matID As Long, ByVal texID As Long)
    If Not Valid(matID) Then Exit Sub
    m_Materials(matID).NormalTextureID = texID
End Sub

Public Sub SetSpecularTexture(ByVal matID As Long, ByVal texID As Long)
    If Not Valid(matID) Then Exit Sub
    m_Materials(matID).SpecularTextureID = texID
End Sub

Public Sub SetLighting(ByVal matID As Long, ByVal value As Boolean)
    If Not Valid(matID) Then Exit Sub
    m_Materials(matID).UseLighting = value
End Sub

Public Sub SetWireframe(ByVal matID As Long, ByVal value As Boolean)
    If Not Valid(matID) Then Exit Sub
    m_Materials(matID).Wireframe = value
End Sub

Public Sub SetDepthTest(ByVal matID As Long, ByVal value As Boolean)
    If Not Valid(matID) Then Exit Sub
    m_Materials(matID).DepthTest = value
End Sub

Public Sub SetShininess(ByVal matID As Long, ByVal value As Single)
    If Not Valid(matID) Then Exit Sub
    m_Materials(matID).Shininess = value
End Sub

' =========================
' APPLY MATERIAL
' =========================

Public Sub Apply(ByVal matID As Long, _
                 ByRef shader As Object, _
                 ByRef TextureSystem As Object, _
                 ByRef lightSystem As Object)

    If Not Valid(matID) Then Exit Sub

    Dim mat As material
    mat = m_Materials(matID)

    ' 1. SHADER BIND
    ' Using the Raw GL call to ensure we bind the specific ID stored in the material
    GL.glUseProgram mat.shaderID

    ' 2. DEPTH TEST
    If mat.DepthTest Then
        GL.glEnable GL_DEPTH_TEST
    Else
        GL.glDisable GL_DEPTH_TEST
    End If

    ' 3. WIREFRAME MODE
    If mat.Wireframe Then
        GL.glPolygonMode GL_FRONT_AND_BACK, GL_LINE
    Else
        GL.glPolygonMode GL_FRONT_AND_BACK, GL_FILL
    End If

    ' 4. TEXTURES
    If mat.UseTextures Then
        If mat.DiffuseTextureID <> -1 Then
            Call TextureSystem.BindTexture(0, mat.DiffuseTextureID)
        End If

        If mat.NormalTextureID <> -1 Then
            Call TextureSystem.BindTexture(1, mat.NormalTextureID)
        End If

        If mat.SpecularTextureID <> -1 Then
            Call TextureSystem.BindTexture(2, mat.SpecularTextureID)
        End If
    End If

    ' 5. LIGHTING & SURFACE UNIFORMS
    If mat.UseLighting Then
        ' PointLightCount/SpotLightCount are properties of the MaterialSystem or LightSystem
        shader.SetUniform1i "uPointLightCount", lightSystem.PointLightCount
        shader.SetUniform1i "uSpotLightCount", lightSystem.SpotLightCount
        
        ' Delegate actual light data binding to the LightSystem
        Call lightSystem.BindLights(shader)
    End If

    ' 6. SURFACE PROPERTIES
    shader.SetUniform1f "uShininess", mat.Shininess
    shader.SetUniform1i "uUseLighting", IIf(mat.UseLighting, 1, 0)
End Sub

' =========================
' HELPERS
' =========================

Public Function GetName(ByVal matID As Long) As String
    If Not Valid(matID) Then Exit Function
    GetName = m_Materials(matID).name
End Function

Public Function count() As Long
    count = m_Count
End Function

Private Function Valid(ByVal matID As Long) As Boolean
    Valid = (matID >= 0 And matID < m_Count)
End Function

Public Sub Clear()
    m_Count = 0
    ReDim m_Materials(0 To 0)
End Sub

Public Sub DebugPrint()
    Debug.Print "---- MATERIAL SYSTEM ----"
    Debug.Print "Materials Count: "; m_Count
    Debug.Print "Global Lights  : "; PointLightCount + SpotLightCount
End Sub




