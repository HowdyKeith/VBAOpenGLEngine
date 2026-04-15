'If you want next upgrades, the REAL physics leap is:

'?? rigidbody impulse solver (stacking + restitution)
'?? capsule-vs-mesh SAT solver (real contact points)
'?? continuous collision detection (CCD sweep + time of impact)
'?? slope friction + walking stability system
'?? physics material blending system

Option Explicit

' ============================================================
' CapsuleCollider.cls
' Static Capsule Collider (VBA ECS v7 Core)
' ============================================================

' =========================
' OWNER
' =========================

Public entityID As Long

' =========================
' SHAPE DATA
' =========================

Public Radius As Single
Public Height As Single

' Capsule axis (default Y-axis)
Public IsVertical As Boolean

' =========================
' TRANSFORM (WORLD SPACE)
' =========================

Public Position As Vector3

' Optional orientation (future support)
Public Rotation As Vector3

' =========================
' PHYSICS SETTINGS
' =========================

Public IsTrigger As Boolean
Public IsStatic As Boolean

Public Mass As Single
Public Bounciness As Single
Public Friction As Single

' =========================
' LAYERING
' =========================

Public Layer As Long

' =========================
' CACHE (OPTIONAL)
' =========================

Public LastCollision As CollisionResult

' =========================
' INIT
' =========================

Private Sub Class_Initialize()

    entityID = -1

    Radius = 0.5
    Height = 2#

    IsVertical = True

    Set Position = New Vector3
    Set Rotation = New Vector3

    IsTrigger = False
    IsStatic = False

    Mass = 1#
    Bounciness = 0#
    Friction = 0.5

    Layer = 0

    Set LastCollision = New CollisionResult

End Sub

' ============================================================
' SET POSITION
' ============================================================

Public Sub SetPosition(ByVal x As Single, ByVal y As Single, ByVal z As Single)

    Position.x = x
    Position.y = y
    Position.z = z

End Sub

' ============================================================
' SET SIZE
' ============================================================

Public Sub SetSize(ByVal Radius As Single, ByVal Height As Single)

    Radius = Radius
    Height = Height

End Sub

' ============================================================
' GET CAPSULE TOP / BOTTOM POINTS (WORLD SPACE)
' ============================================================

Public Sub GetEndpoints(ByRef top As Vector3, ByRef bottom As Vector3)

    Dim halfHeight As Single
    halfHeight = (Height * 0.5) - Radius

    Set top = New Vector3
    Set bottom = New Vector3

    top.x = Position.x
    top.y = Position.y + halfHeight
    top.z = Position.z

    bottom.x = Position.x
    bottom.y = Position.y - halfHeight
    bottom.z = Position.z

End Sub

' ============================================================
' SIMPLE BOUNDS (FOR BROADPHASE)
' ============================================================

Public Sub GetAABB(ByRef minV As Vector3, ByRef maxV As Vector3)

    Dim halfH As Single
    halfH = (Height * 0.5)

    Set minV = New Vector3
    Set maxV = New Vector3

    minV.x = Position.x - Radius
    minV.y = Position.y - halfH
    minV.z = Position.z - Radius

    maxV.x = Position.x + Radius
    maxV.y = Position.y + halfH
    maxV.z = Position.z + Radius

End Sub

' ============================================================
' APPLY COLLISION RESULT (FROM PHYSICS ENGINE)
' ============================================================

Public Sub ApplyCollision(ByRef result As CollisionResult)

    If result Is Nothing Then Exit Sub
    If Not result.hit Then Exit Sub

    LastCollision.CopyFrom result

    ' Basic penetration correction
    Position.x = Position.x + result.Correction.x
    Position.y = Position.y + result.Correction.y
    Position.z = Position.z + result.Correction.z

End Sub

' ============================================================
' DEBUG
' ============================================================

Public Sub DebugPrint()

    Debug.Print "---- CAPSULE COLLIDER ----"
    Debug.Print "Entity: "; entityID
    Debug.Print "Pos: "; Position.x; Position.y; Position.z
    Debug.Print "Radius: "; Radius; " Height: "; Height
    Debug.Print "Mass: "; Mass

End Sub




