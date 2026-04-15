'If you want next upgrades, the REAL physics leap is:

'??
'full manifold generation (contact points per pair)
'??
'impulse solver (stacking rigid bodies)
'??
'friction + restitution system
'??
'continuous collision detection (CCD sweep)
'??
'debug collision visualizer (gizmos)
Option Explicit

' ============================================================
' CollisionPair.cls
' Physics Broadphase / Collision Pair Container (v7 Core)
' ============================================================

' =========================
' ENTITY REFERENCES
' =========================

Public EntityA As Long
Public EntityB As Long

' =========================
' COLLIDERS (OPTIONAL OBJECT REFERENCES)
' =========================

Public ColliderA As Object
Public ColliderB As Object

' =========================
' PAIR STATE
' =========================

Public IsActive As Boolean
Public IsTrigger As Boolean

' =========================
' COLLISION FLAGS
' =========================

Public WasColliding As Boolean
Public IsColliding As Boolean

' =========================
' CONTACT DATA (LIGHTWEIGHT CACHE)
' =========================

Public ContactPoint As Vector3
Public ContactNormal As Vector3

Public Penetration As Single
Public distance As Single

' =========================
' RESPONSE DATA
' =========================

Public ImpulseA As Vector3
Public ImpulseB As Vector3

Public CorrectionA As Vector3
Public CorrectionB As Vector3

' =========================
' INTERNAL IDS
' =========================

Public PairID As Long
Public LastFrameUpdated As Long

' =========================
' INIT
' =========================

Private Sub Class_Initialize()

    EntityA = -1
    EntityB = -1

    Set ColliderA = Nothing
    Set ColliderB = Nothing

    IsActive = True
    IsTrigger = False

    WasColliding = False
    IsColliding = False

    Set ContactPoint = New Vector3
    Set ContactNormal = New Vector3

    Set ImpulseA = New Vector3
    Set ImpulseB = New Vector3

    Set CorrectionA = New Vector3
    Set CorrectionB = New Vector3

    Penetration = 0
    distance = 0

    PairID = -1
    LastFrameUpdated = 0

End Sub

' ============================================================
' INITIALIZE PAIR
' ============================================================

Public Sub Init(ByVal a As Long, ByVal b As Long)

    EntityA = a
    EntityB = b

    IsActive = True

    WasColliding = False
    IsColliding = False

End Sub

' ============================================================
' UPDATE COLLISION STATE
' ============================================================

Public Sub SetCollisionState(ByVal hit As Boolean)

    WasColliding = IsColliding
    IsColliding = hit

End Sub

' ============================================================
' RESET CONTACT DATA
' ============================================================

Public Sub ClearContacts()

    ContactPoint.x = 0
    ContactPoint.y = 0
    ContactPoint.z = 0

    ContactNormal.x = 0
    ContactNormal.y = 0
    ContactNormal.z = 0

    Penetration = 0
    distance = 0

    ImpulseA.x = 0
    ImpulseA.y = 0
    ImpulseA.z = 0

    ImpulseB.x = 0
    ImpulseB.y = 0
    ImpulseB.z = 0

    CorrectionA.x = 0
    CorrectionA.y = 0
    CorrectionA.z = 0

    CorrectionB.x = 0
    CorrectionB.y = 0
    CorrectionB.z = 0

End Sub

' ============================================================
' SWAP ENTITIES (NORMALIZATION)
' ============================================================

Public Sub Normalize()

    If EntityA <= EntityB Then Exit Sub

    Dim temp As Long
    temp = EntityA
    EntityA = EntityB
    EntityB = temp

    Dim tempObj As Object

    Set tempObj = ColliderA
    Set ColliderA = ColliderB
    Set ColliderB = tempObj

End Sub

' ============================================================
' HASH FOR PAIR COMPARISON (FAST LOOKUP KEY)
' ============================================================

Public Function GetHash() As Long

    Dim a As Long
    Dim b As Long

    a = EntityA
    b = EntityB

    If a > b Then
        Dim t As Long
        t = a
        a = b
        b = t
    End If

    GetHash = (a * 73856093) Xor (b * 19349663)

End Function

' ============================================================
' DEBUG
' ============================================================

Public Sub DebugPrint()

    Debug.Print "---- COLLISION PAIR ----"
    Debug.Print "A: "; EntityA; " B: "; EntityB
    Debug.Print "Active: "; IsActive
    Debug.Print "Colliding: "; IsColliding
    Debug.Print "Penetration: "; Penetration

End Sub



