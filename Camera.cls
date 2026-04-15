Option Explicit

' ============================================================
' Class      : Camera
' Version    : v7.2
' Description: Core Camera Component for VBA OpenGL.
'              Synchronized with GL.bas v8.51 Facade.
' ============================================================

' =========================
' IDENTITY & TRANSFORM
' =========================
Public id As Long
Public name As String
Public Active As Boolean

Public Position As Vector3
Public Rotation As Vector3

' =========================
' CAMERA SETTINGS
' =========================
Public FOV As Single
Public NearPlane As Single
Public FarPlane As Single
Public AspectRatio As Single

' =========================
' INIT
' =========================
Private Sub Class_Initialize()
    id = -1
    name = "MainCamera"
    Active = True

    Set Position = New Vector3
    Set Rotation = New Vector3

    ' Standard Defaults
    FOV = 60#
    NearPlane = 0.1
    FarPlane = 1000#
    AspectRatio = 1.7777 ' 16:9
End Sub

' ============================================================
' UPDATE HOOK
' ============================================================
Public Sub Update(ByRef inputMgr As Object, ByVal dt As Single)
    ' Placeholder for smoothing, shake, or lerping logic
End Sub

' ============================================================
' VIEW MATRIX APPLICATION
' ============================================================
''' <summary>
''' Applies the Camera's View Transform (Rotation then Translation).
''' </summary>
Public Sub ApplyView()
    ' 1. Reset Modelview Matrix
    GL.glMatrixMode GL.GL_MODELVIEW
    GL.glLoadIdentity

    ' 2. Apply Rotations (Inverted for Camera Space)
    ' Order: Pitch (X), Yaw (Y), Roll (Z)
    GL.glRotatef -Rotation.x, 1, 0, 0
    GL.glRotatef -Rotation.y, 0, 1, 0
    GL.glRotatef -Rotation.z, 0, 0, 1

    ' 3. Apply Translation (Inverted for Camera Space)
    GL.glTranslatef -Position.x, -Position.y, -Position.z
End Sub

' ============================================================
' PROJECTION MATRIX APPLICATION
' ============================================================
''' <summary>
''' Applies the Perspective Projection.
''' </summary>
Public Sub ApplyProjection()
    GL.glMatrixMode GL.GL_PROJECTION
    GL.glLoadIdentity

    ' Perspective calculation via GLU bridge
    GL.gluPerspective FOV, AspectRatio, NearPlane, FarPlane

    ' Return context to Modelview
    GL.glMatrixMode GL.GL_MODELVIEW
End Sub

' ============================================================
' UTILITIES
' ============================================================

''' <summary>
''' Forces the camera to face a specific world coordinate.
''' </summary>
Public Sub LookAt(ByVal tx As Single, ByVal ty As Single, ByVal tz As Single)
    Dim dx As Single, dy As Single, dz As Single
    
    dx = tx - Position.x
    dy = ty - Position.y
    dz = tz - Position.z

    Rotation.y = Atn2(dx, dz) * 180# / 3.14159265
    Rotation.x = -Atn2(dy, Sqr(dx * dx + dz * dz)) * 180# / 3.14159265
End Sub

Public Sub SetPosition(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Position.x = x: Position.y = y: Position.z = z
End Sub

Public Sub SetRotation(ByVal pitch As Single, ByVal yaw As Single, ByVal roll As Single)
    Rotation.x = pitch: Rotation.y = yaw: Rotation.z = roll
End Sub

' ============================================================
' VECTOR CALCULATIONS
' ============================================================

Public Function GetForward() As Vector3
    Dim v As New Vector3
    Dim pitchRad As Single: pitchRad = DegToRad(Rotation.x)
    Dim yawRad As Single: yawRad = DegToRad(Rotation.y)

    v.x = Cos(yawRad) * Cos(pitchRad)
    v.y = Sin(pitchRad)
    v.z = Sin(yawRad) * Cos(pitchRad)
    Set GetForward = v
End Function

Public Function GetRight() As Vector3
    Dim v As New Vector3
    Dim yawRad As Single: yawRad = DegToRad(Rotation.y - 90)
    
    v.x = Sin(yawRad)
    v.y = 0
    v.z = Cos(yawRad)
    Set GetRight = v
End Function

' ============================================================
' MATH HELPERS
' ============================================================

Private Function DegToRad(ByVal deg As Single) As Single
    DegToRad = deg * 3.14159265 / 180#
End Function

Private Function Atn2(ByVal y As Single, ByVal x As Single) As Single
    If x = 0 Then
        If y > 0 Then Atn2 = 1.570796 Else Atn2 = -1.570796
    Else
        Atn2 = Atn(y / x)
        If x < 0 Then Atn2 = Atn2 + 3.14159265
    End If
End Function

' ============================================================
' CLEANUP
' ============================================================
Private Sub Class_Terminate()
    Set Position = Nothing
    Set Rotation = Nothing
End Sub


