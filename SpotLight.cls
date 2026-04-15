Option Explicit

' ============================================================
' SpotLight.cls
' ECS Spot Light Component (VBA OpenGL v7 Engine Core)
' ============================================================

' =========================
' TRANSFORM
' =========================

Public Position As Vector3
Public direction As Vector3

' =========================
' LIGHT PROPERTIES
' =========================

Public Color As Vector3          ' RGB (0..1)
Public Intensity As Single

' =========================
' SPOT CONE SETTINGS
' =========================

Public InnerCutoff As Single     ' radians or cos(angle)
Public OuterCutoff As Single     ' radians or cos(angle)

' =========================
' ATTENUATION
' =========================

Public Constant As Single
Public Linear As Single
Public Quadratic As Single

' =========================
' STATE
' =========================

Public Enabled As Boolean
Public CastShadows As Boolean

Public Range As Single

' =========================
' INTERNAL ID (for ECS / shader binding)
' =========================

Public LightID As Long

' =========================
' INIT
' =========================

Private Sub Class_Initialize()

    Set Position = New Vector3
    Set direction = New Vector3
    Set Color = New Vector3

    ' default white light
    Color.x = 1
    Color.y = 1
    Color.z = 1

    Intensity = 1

    ' default cone (about 30–45 degrees)
    InnerCutoff = 0.9
    OuterCutoff = 0.85

    ' physically plausible defaults
    Constant = 1#
    Linear = 0.09
    Quadratic = 0.032

    Range = 25

    Enabled = True
    CastShadows = False

    LightID = -1

End Sub

' =========================
' UPDATE DIRECTION (LOOK AT TARGET)
' =========================

Public Sub LookAt(ByVal target As Vector3)

    direction.x = target.x - Position.x
    direction.y = target.y - Position.y
    direction.z = target.z - Position.z

    Call Normalize(direction)

End Sub

' =========================
' NORMALIZE VECTOR
' =========================

Private Sub Normalize(ByRef v As Vector3)

    Dim len As Single

    len = Sqr(v.X * v.X + v.Y * v.Y + v.Z * v.Z)

    If len = 0 Then Exit Sub

    v.X = v.X / len
    v.Y = v.Y / len
    v.Z = v.Z / len

End Sub

' =========================
' APPLY TO SHADER
' =========================

Public Sub ApplyToShader(ByVal programID As LongPtr, ByVal index As Long)

    If Not Enabled Then Exit Sub

    modGL.g_ActiveProgram = programID
    Dim prefix As String
    prefix = "spotLights[" & CStr(index) & "]"

    Call SetShaderVec3(prefix & ".position", Position.x, Position.y, Position.z)
    Call SetShaderVec3(prefix & ".direction", direction.x, direction.y, direction.z)

    Call SetShaderVec3(prefix & ".color", Color.x, Color.y, Color.z)

    Call SetShaderFloat(prefix & ".intensity", Intensity)

    Call SetShaderFloat(prefix & ".innerCutoff", InnerCutoff)
    Call SetShaderFloat(prefix & ".outerCutoff", OuterCutoff)

    Call SetShaderFloat(prefix & ".constant", Constant)
    Call SetShaderFloat(prefix & ".linear", Linear)
    Call SetShaderFloat(prefix & ".quadratic", Quadratic)

    Call SetShaderFloat(prefix & ".range", Range)

End Sub

' =========================
' DEBUG HELPERS
' =========================

Public Sub SetColor(ByVal r As Single, ByVal g As Single, ByVal b As Single)

    Color.x = r
    Color.y = g
    Color.z = b

End Sub

Public Sub SetDirection(ByVal x As Single, ByVal y As Single, ByVal z As Single)

    direction.x = x
    direction.y = y
    direction.z = z

    Call Normalize(direction)

End Sub

' =========================
' TOGGLE
' =========================

Public Sub Enable()
    Enabled = True
End Sub

Public Sub Disable()
    Enabled = False
End Sub



