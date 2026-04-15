'If you want next upgrades, the real “engine leap” is:

'?? capsule collider + real collision response
'?? slope handling + step offset system
'?? camera bobbing + head sway
'?? sprint + stamina system
'?? full character controller (Unity-style)
Option Explicit

' ============================================================
' FPSController.cls
' First Person Player Controller (VBA ECS v7 Core)
' ============================================================

' =========================
' REFERENCES
' =========================

Public Camera As GLCamera
'FIX
Public InputSys As InputManager
Public physics As PhysicsSystem   ' optional hook

' =========================
' STATE
' =========================

Public Position As Vector3
Public Velocity As Vector3

' =========================
' SETTINGS
' =========================

Public MoveSpeed As Single
Public MouseSensitivity As Single

Public JumpStrength As Single
Public Gravity As Single

Public IsGrounded As Boolean

' =========================
' INIT
' =========================

Private Sub Class_Initialize()

    Set Camera = New GLCamera
    Set InputSys = Nothing
    Set physics = Nothing

    Set Position = New Vector3
    Set Velocity = New Vector3

    MoveSpeed = 6#
    MouseSensitivity = 0.15
    JumpStrength = 5#
    Gravity = -9.81

    IsGrounded = False

End Sub

' ============================================================
' UPDATE (CALL EACH FRAME)
' ============================================================

Public Sub Update(ByVal dt As Single)

    If InputSys Is Nothing Then Exit Sub
    If Camera Is Nothing Then Exit Sub

    Call UpdateLook
    Call UpdateMovement(dt)
    Call ApplyPhysics(dt)

    ' sync camera
    Camera.Position.x = Position.x
    Camera.Position.y = Position.y
    Camera.Position.z = Position.z

End Sub

' ============================================================
' MOUSE LOOK
' ============================================================

Private Sub UpdateLook()

    Camera.Rotation.y = Camera.Rotation.y + InputSys.MouseDeltaX() * MouseSensitivity
    Camera.Rotation.x = Camera.Rotation.x + InputSys.MouseDeltaY() * MouseSensitivity

    If Camera.Rotation.x > 89 Then Camera.Rotation.x = 89
    If Camera.Rotation.x < -89 Then Camera.Rotation.x = -89

End Sub

' ============================================================
' MOVEMENT (WASD)
' ============================================================

Private Sub UpdateMovement(ByVal dt As Single)

    Dim forward As Vector3
    Dim right As Vector3

    Set forward = Camera.GetForward()
    Set right = Camera.GetRight()

    Dim moveX As Single
    Dim moveZ As Single

    moveX = 0
    moveZ = 0

    ' -------------------------
    ' INPUT
    ' -------------------------

    If InputSys.KeyDown(87) Then moveZ = moveZ + 1 ' W
    If InputSys.KeyDown(83) Then moveZ = moveZ - 1 ' S
    If InputSys.KeyDown(65) Then moveX = moveX - 1 ' A
    If InputSys.KeyDown(68) Then moveX = moveX + 1 ' D

    ' -------------------------
    ' APPLY DIRECTION
    ' -------------------------

    Position.x = Position.x + (forward.x * moveZ + right.x * moveX) * MoveSpeed * dt
    Position.z = Position.z + (forward.z * moveZ + right.z * moveX) * MoveSpeed * dt

    ' -------------------------
    ' JUMP
    ' -------------------------

    If InputSys.KeyPressed(32) And IsGrounded Then ' SPACE
        Velocity.y = JumpStrength
        IsGrounded = False
    End If

End Sub

' ============================================================
' PHYSICS (GRAVITY + OPTIONAL WORLD COLLISION)
' ============================================================

Private Sub ApplyPhysics(ByVal dt As Single)
    Velocity.y = Velocity.y + Gravity * dt
    Position.y = Position.y + Velocity.y * dt

    ' 1. Basic Hard-Floor (Safety Net)
    If Position.y <= 0 Then
        Position.y = 0
        Velocity.y = 0
        IsGrounded = True
    End If

    ' 2. Physics Hook (The Raycast fix)
    If Not physics Is Nothing Then
        Dim result As IntersectionResult
        Set result = New IntersectionResult ' Reset result carrier

        ' Origin is the player's current position
        physics.RaycastDown Position, result

        If result.hit Then
            ' Check if we are "inside" or "below" the detected point
            ' We add a tiny offset (0.01) to keep the player slightly above
            If Position.y < result.point.y + 0.01 Then
                Position.y = result.point.y
                Velocity.y = 0
                IsGrounded = True
            End If
        End If
    End If
End Sub

' ============================================================
' SET INPUT SYSTEM
' ============================================================

Public Sub BindInput(ByRef InputSystem As InputManager)

    Set InputSys = InputSystem

End Sub

' ============================================================
' SET PHYSICS WORLD (OPTIONAL)
' ============================================================

Public Sub BindPhysics(ByRef physicsWorld As PhysicsSystem)

    Set physics = physicsWorld

End Sub

' ============================================================
' RESET PLAYER
' ============================================================

Public Sub Reset()

    Position.x = 0
    Position.y = 0
    Position.z = 0

    Velocity.x = 0
    Velocity.y = 0
    Velocity.z = 0

    IsGrounded = False

End Sub

' ============================================================
' DEBUG
' ============================================================

Public Sub DebugPrint()

    Debug.Print "---- FPS CONTROLLER ----"
    Debug.Print "Pos: "; Position.x; Position.y; Position.z
    Debug.Print "VelY: "; Velocity.y
    Debug.Print "Grounded: "; IsGrounded

End Sub



