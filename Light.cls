Option Explicit

' ============================================================
' Lighting.cls
' Per-Entity Light Component (VBA ECS v7 Engine Core)
' ============================================================
'If you want next upgrades, the real “engine leap” is:
'shadow mapping system
'clustered forward lighting
'GPU light culling
'deferred lighting(g - Buffer)
'volumetric lighting
' =========================
' IDENTITY
' =========================

Public id As Long
Public entityID As Long
Public name As String

Public Active As Boolean

' =========================
' LIGHT TYPE
' =========================

Public Enum LightType
    ltPoint = 0
    ltSpot = 1
    ltDirectional = 2
End Enum

Public TypeOfLight As LightType

' =========================
' TRANSFORM
' =========================

Public Position As Vector3
Public direction As Vector3

' =========================
' COLOR + INTENSITY
' =========================

Public Color As Vector3
Public Intensity As Single

' =========================
' RANGE / FALLOFF
' =========================

Public Range As Single

' =========================
' SPOTLIGHT PARAMETERS
' =========================

Public Cutoff As Single
Public OuterCutoff As Single

' =========================
' SHADOW FLAGS (FUTURE EXTENSION)
' =========================

Public CastShadows As Boolean

' =========================
' INIT
' =========================

Private Sub Class_Initialize()

    id = -1
    entityID = -1
    name = ""

    Active = True

    TypeOfLight = ltPoint

    Set Position = New Vector3
    Set direction = New Vector3
    Set Color = New Vector3

    Color.x = 1
    Color.y = 1
    Color.z = 1

    Intensity = 1
    Range = 10

    Cutoff = 12.5
    OuterCutoff = 17.5

    CastShadows = False

End Sub

' =========================
' CONFIG HELPERS
' =========================

Public Sub SetPointLight()

    TypeOfLight = ltPoint

End Sub

Public Sub SetSpotLight()

    TypeOfLight = ltSpot

End Sub

Public Sub SetDirectionalLight()

    TypeOfLight = ltDirectional

End Sub

' =========================
' POSITION / DIRECTION
' =========================

Public Sub SetPosition(ByVal x As Single, ByVal y As Single, ByVal z As Single)

    Position.x = x
    Position.y = y
    Position.z = z

End Sub

Public Sub SetDirection(ByVal x As Single, ByVal y As Single, ByVal z As Single)

    direction.x = x
    direction.y = y
    direction.z = z

End Sub

' =========================
' COLOR CONTROL
' =========================

Public Sub SetColor(ByVal r As Single, ByVal g As Single, ByVal b As Single)

    Color.x = r
    Color.y = g
    Color.z = b

End Sub

Public Sub SetIntensity(ByVal value As Single)

    Intensity = value

End Sub

' =========================
' SPOT SETTINGS
' =========================

Public Sub SetSpotAngles(ByVal inner As Single, ByVal outer As Single)

    Cutoff = inner
    OuterCutoff = outer

End Sub

' =========================
' ENABLE / DISABLE
' =========================

Public Sub Enable()
    Active = True
End Sub

Public Sub Disable()
    Active = False
End Sub

' =========================
' DEBUG
' =========================

Public Sub DebugPrint()

    Debug.Print "---- LIGHT ----"
    Debug.Print "ID: "; id; " Entity: "; entityID
    Debug.Print "Type: "; TypeOfLight
    Debug.Print "Pos: "; Position.x; Position.y; Position.z
    Debug.Print "Intensity: "; Intensity

End Sub



