'If you want next upgrades, the REAL physics leap is:

'?? manifold collision system (edge/face contact sets)
'?? continuous collision detection (CCD)
'?? impulse solver (real rigid body stacking)
'?? friction + restitution system
'?? sleep/activation system
Option Explicit

' ============================================================
' CollisionResult.cls
' Physics Collision / Raycast Result (VBA ECS v7 Core)
' ============================================================

' =========================
' HIT STATE
' =========================

Public hit As Boolean

' =========================
' IMPACT DATA
' =========================

Public point As Vector3
Public normal As Vector3

Public distance As Single

' =========================
' OBJECT REFERENCES
' =========================

Public EntityA As Long
Public EntityB As Long

' Optional pointers (if using object-based physics)
Public ColliderA As Object
Public ColliderB As Object

' =========================
' RESPONSE DATA (FOR RESOLUTION)
' =========================

Public PenetrationDepth As Single
Public Correction As Vector3

Public Impulse As Vector3

' =========================
' METADATA
' =========================

Public TriangleIndex As Long
Public MaterialID As Long

' =========================
' INIT
' =========================

Private Sub Class_Initialize()

    hit = False

    Set point = New Vector3
    Set normal = New Vector3
    Set Correction = New Vector3
    Set Impulse = New Vector3

    distance = 0
    PenetrationDepth = 0

    EntityA = -1
    EntityB = -1

    ColliderA = Nothing
    ColliderB = Nothing

    TriangleIndex = -1
    MaterialID = -1

End Sub

' ============================================================
' RESET RESULT (REUSE FOR PERFORMANCE)
' ============================================================

Public Sub Clear()

    hit = False

    point.x = 0
    point.y = 0
    point.z = 0

    normal.x = 0
    normal.y = 0
    normal.z = 0

    Correction.x = 0
    Correction.y = 0
    Correction.z = 0

    Impulse.x = 0
    Impulse.y = 0
    Impulse.z = 0

    distance = 0
    PenetrationDepth = 0

    EntityA = -1
    EntityB = -1

    ColliderA = Nothing
    ColliderB = Nothing

    TriangleIndex = -1
    MaterialID = -1

End Sub

' ============================================================
' COPY RESULT (FOR SAFETY / QUEUING)
' ============================================================

Public Sub CopyFrom(ByRef src As CollisionResult)

    If src Is Nothing Then Exit Sub

    hit = src.hit

    point.x = src.point.x
    point.y = src.point.y
    point.z = src.point.z

    normal.x = src.normal.x
    normal.y = src.normal.y
    normal.z = src.normal.z

    distance = src.distance
    PenetrationDepth = src.PenetrationDepth

    EntityA = src.EntityA
    EntityB = src.EntityB

    TriangleIndex = src.TriangleIndex
    MaterialID = src.MaterialID

    Set ColliderA = src.ColliderA
    Set ColliderB = src.ColliderB

End Sub

' ============================================================
' DEBUG
' ============================================================

Public Sub DebugPrint()

    Debug.Print "---- COLLISION RESULT ----"
    Debug.Print "Hit: "; hit
    Debug.Print "Point: "; point.x; point.y; point.z
    Debug.Print "Normal: "; normal.x; normal.y; normal.z
    Debug.Print "Distance: "; distance
    Debug.Print "Penetration: "; PenetrationDepth

End Sub



