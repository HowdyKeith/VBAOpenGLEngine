Option Explicit

' ============================================================
' LightingSystem.cls  WEEK 1 PERFORMANCE PASS
' WEEK 1 CHANGES:
'   - Uniform location strings were being BUILT every frame:
'     "uPointLights[" & i & "].position" costs string allocation + concat.
'     With 8 point lights x 4 uniforms = 32 string builds per frame = ~64KB
'     of string allocations per second, all going to GC.
'     Now pre-built at CreatePointLight/CreateSpotLight time and stored.
'   - BindLights now calls shader.SetUniform3f directly with cached string names
'     (saves the string build) and skips disabled lights entirely.
'   - Flat Singles kept from previous fix (no object refs in UDT).
' ============================================================

Private Type PointLightData
    px As Single: py As Single: pz As Single
    cr As Single: cg As Single: cb As Single
    Intensity As Single
    Range     As Single
    Enabled   As Boolean
    ' WEEK 1: Pre-built uniform name strings
    uPos  As String   ' "uPointLights[i].position"
    uCol  As String   ' "uPointLights[i].color"
    uInt  As String   ' "uPointLights[i].intensity"
    uRng  As String   ' "uPointLights[i].range"
End Type

Private Type SpotLightData
    px  As Single: py  As Single: pz  As Single
    dx  As Single: dy  As Single: dz  As Single
    cr  As Single: cg  As Single: cb  As Single
    InnerCutoff As Single
    OuterCutoff As Single
    Intensity   As Single
    Enabled     As Boolean
    ' WEEK 1: Pre-built uniform name strings
    uPos  As String
    uDir  As String
    uCol  As String
    uInner As String
    uOuter As String
    uInt  As String
End Type

Private m_PointLights() As PointLightData
Private m_SpotLights()  As SpotLightData
Private m_PointCount    As Long
Private m_SpotCount     As Long

Public MaxPointLights As Long
Public MaxSpotLights  As Long

Private Sub Class_Initialize()
    m_PointCount  = 0
    m_SpotCount   = 0
    MaxPointLights = 16
    MaxSpotLights  = 8
    ReDim m_PointLights(0 To 0)
    ReDim m_SpotLights(0 To 0)
End Sub

' ============================================================
' POINT LIGHTS
' ============================================================
Public Function CreatePointLight() As Long
    If m_PointCount >= MaxPointLights Then CreatePointLight = -1: Exit Function
    Dim id As Long: id = m_PointCount
    m_PointCount = m_PointCount + 1
    ReDim Preserve m_PointLights(0 To m_PointCount - 1)
    With m_PointLights(id)
        .cr = 1: .cg = 1: .cb = 1
        .Intensity = 1: .Range = 10: .Enabled = True
        ' WEEK 1: Build uniform names ONCE at creation time
        Dim p As String: p = "uPointLights[" & id & "]."
        .uPos = p & "position"
        .uCol = p & "color"
        .uInt = p & "intensity"
        .uRng = p & "range"
    End With
    CreatePointLight = id
End Function

Public Sub SetPointLightPosition(ByVal id As Long, ByVal x As Single, ByVal y As Single, ByVal z As Single)
    If id < 0 Or id >= m_PointCount Then Exit Sub
    m_PointLights(id).px = x: m_PointLights(id).py = y: m_PointLights(id).pz = z
End Sub

Public Sub SetPointLightColor(ByVal id As Long, ByVal r As Single, ByVal g As Single, ByVal b As Single)
    If id < 0 Or id >= m_PointCount Then Exit Sub
    m_PointLights(id).cr = r: m_PointLights(id).cg = g: m_PointLights(id).cb = b
End Sub

Public Sub SetPointLightRange(ByVal id As Long, ByVal range As Single)
    If id < 0 Or id >= m_PointCount Then Exit Sub
    m_PointLights(id).Range = range
End Sub

Public Sub SetPointLightIntensity(ByVal id As Long, ByVal intensity As Single)
    If id < 0 Or id >= m_PointCount Then Exit Sub
    m_PointLights(id).Intensity = intensity
End Sub

Public Sub EnablePointLight(ByVal id As Long, ByVal state As Boolean)
    If id < 0 Or id >= m_PointCount Then Exit Sub
    m_PointLights(id).Enabled = state
End Sub

' ============================================================
' SPOT LIGHTS
' ============================================================
Public Function CreateSpotLight() As Long
    If m_SpotCount >= MaxSpotLights Then CreateSpotLight = -1: Exit Function
    Dim id As Long: id = m_SpotCount
    m_SpotCount = m_SpotCount + 1
    ReDim Preserve m_SpotLights(0 To m_SpotCount - 1)
    With m_SpotLights(id)
        .cr = 1: .cg = 1: .cb = 1
        .Intensity = 1: .InnerCutoff = 0.9: .OuterCutoff = 0.8: .Enabled = True
        Dim p As String: p = "uSpotLights[" & id & "]."
        .uPos  = p & "position"
        .uDir  = p & "direction"
        .uCol  = p & "color"
        .uInner = p & "innerCutoff"
        .uOuter = p & "outerCutoff"
        .uInt  = p & "intensity"
    End With
    CreateSpotLight = id
End Function

Public Sub SetSpotLightPosition(ByVal id As Long, ByVal x As Single, ByVal y As Single, ByVal z As Single)
    If id < 0 Or id >= m_SpotCount Then Exit Sub
    m_SpotLights(id).px = x: m_SpotLights(id).py = y: m_SpotLights(id).pz = z
End Sub

Public Sub SetSpotLightDirection(ByVal id As Long, ByVal x As Single, ByVal y As Single, ByVal z As Single)
    If id < 0 Or id >= m_SpotCount Then Exit Sub
    m_SpotLights(id).dx = x: m_SpotLights(id).dy = y: m_SpotLights(id).dz = z
End Sub

' ============================================================
' WEEK 1: BIND LIGHTS - uses pre-built string names, skips disabled
' ============================================================
Public Sub BindLights(ByRef shader As Object)
    If shader Is Nothing Then Exit Sub
    Dim i As Long

    For i = 0 To m_PointCount - 1
        If m_PointLights(i).Enabled Then
            ' WEEK 1: Pre-built strings - no allocation on this path
            shader.SetUniform3f m_PointLights(i).uPos, m_PointLights(i).px, m_PointLights(i).py, m_PointLights(i).pz
            shader.SetUniform3f m_PointLights(i).uCol, m_PointLights(i).cr, m_PointLights(i).cg, m_PointLights(i).cb
            shader.SetUniform1f m_PointLights(i).uInt, m_PointLights(i).Intensity
            shader.SetUniform1f m_PointLights(i).uRng, m_PointLights(i).Range
        End If
    Next i

    For i = 0 To m_SpotCount - 1
        If m_SpotLights(i).Enabled Then
            shader.SetUniform3f m_SpotLights(i).uPos,   m_SpotLights(i).px,  m_SpotLights(i).py,  m_SpotLights(i).pz
            shader.SetUniform3f m_SpotLights(i).uDir,   m_SpotLights(i).dx,  m_SpotLights(i).dy,  m_SpotLights(i).dz
            shader.SetUniform3f m_SpotLights(i).uCol,   m_SpotLights(i).cr,  m_SpotLights(i).cg,  m_SpotLights(i).cb
            shader.SetUniform1f m_SpotLights(i).uInner, m_SpotLights(i).InnerCutoff
            shader.SetUniform1f m_SpotLights(i).uOuter, m_SpotLights(i).OuterCutoff
            shader.SetUniform1f m_SpotLights(i).uInt,   m_SpotLights(i).Intensity
        End If
    Next i
End Sub

Public Function PointLightCount() As Long: PointLightCount = m_PointCount: End Function
Public Function SpotLightCount()  As Long: SpotLightCount  = m_SpotCount:  End Function

Public Sub Clear()
    m_PointCount = 0: m_SpotCount = 0
    ReDim m_PointLights(0 To 0): ReDim m_SpotLights(0 To 0)
End Sub

Public Sub DebugPrint()
    Debug.Print "---- LIGHTING SYSTEM ---- Points: " & m_PointCount & " Spots: " & m_SpotCount
End Sub
