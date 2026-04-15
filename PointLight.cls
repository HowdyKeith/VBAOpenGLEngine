Option Explicit

' ============================================================
' PointLight.cls
' ECS Point Light Component (VBA OpenGL v7 Engine Core)
' ============================================================

' =========================
' TRANSFORM
' =========================

Public Position As Vector3

' =========================
' LIGHT PROPERTIES
' =========================

Public Color As Vector3
Public Intensity As Single

' =========================
' ATTENUATION
' =========================

Public Constant As Single
Public Linear As Single
Public Quadratic As Single

' =========================
' RANGE / CULLING
' =========================

Public Range As Single

' =========================
' STATE
' =========================

Public Enabled As Boolean
Public CastShadows As Boolean

' =========================
' ECS ID
' =========================

Public LightID As Long

' =========================
' INIT
' =========================

Private Sub Class_Initialize()

    Set Position = New Vector3
    Set Color = New Vector3

    ' default white light
    Color.x = 1
    Color.y = 1
    Color.z = 1

    Intensity = 1#

    ' physically plausible defaults
    Constant = 1#
    Linear = 0.09
    Quadratic = 0.032

    Range = 20#

    Enabled = True
    CastShadows = False

    LightID = -1

End Sub

' =========================
' APPLY TO SHADER
' =========================

Public Sub ApplyToShader(ByVal programID As LongPtr, ByVal index As Long)

    If Not Enabled Then Exit Sub

    modGL.g_ActiveProgram = programID
    Dim prefix As String
    prefix = "pointLights[" & CStr(index) & "]"

    Call SetShaderVec3(prefix & ".position", Position.x, Position.y, Position.z)

    Call SetShaderVec3(prefix & ".color", Color.x, Color.y, Color.z)

    Call SetShaderFloat(prefix & ".intensity", Intensity)

    Call SetShaderFloat(prefix & ".constant", Constant)
    Call SetShaderFloat(prefix & ".linear", Linear)
    Call SetShaderFloat(prefix & ".quadratic", Quadratic)

    Call SetShaderFloat(prefix & ".range", Range)

End Sub

' =========================
' SET COLOR
' =========================

Public Sub SetColor(ByVal r As Single, ByVal g As Single, ByVal b As Single)

    Color.x = r
    Color.y = g
    Color.z = b

End Sub

' =========================
' MOVE LIGHT
' =========================

Public Sub SetPosition(ByVal x As Single, ByVal y As Single, ByVal z As Single)

    Position.x = x
    Position.y = y
    Position.z = z

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

' =========================
' DEBUG
' =========================

Public Sub DebugPrint()

    Debug.Print "---- POINT LIGHT ----"
    Debug.Print "Pos: "; Position.x; Position.y; Position.z
    Debug.Print "Color: "; Color.x; Color.y; Color.z
    Debug.Print "Intensity: "; Intensity

End Sub




