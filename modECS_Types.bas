Option Explicit

' =========================================================
' ECS v8 CORE TYPES MODULE
' No class-level Type declarations allowed elsewhere
' =========================================================

' =========================
' BVH NODE
' =========================
Public Type BVHNode
    minX As Single
    minY As Single
    minZ As Single

    maxX As Single
    maxY As Single
    maxZ As Single

    left As Long
    right As Long
    parent As Long

    entityID As Long
    isLeaf As Boolean
End Type

' =========================
' COLLISION RESULT
' =========================
Public Type CollisionHit
    hit As Boolean
    distance As Single

    nx As Single
    ny As Single
    nz As Single

    px As Single
    py As Single
    pz As Single

    entityID As Long
End Type

' =========================
' SPATIAL GRID CELL
' =========================
Public Type GridCell
    x As Long
    y As Long
    z As Long
End Type


