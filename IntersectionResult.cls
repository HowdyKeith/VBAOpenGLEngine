Option Explicit

' ============================================================
' Class      : IntersectionResult
' Version    : v1.0
' Description: Result carrier for raycasts and intersection tests.
'              Used by PhysicsSystem.RaycastDown and FPSController.
' NOTE: This class was MISSING from the project - FPSController.cls and
'       PhysicsSystem.cls both reference it but it was never defined.
'       CollisionResult.cls covers rigid-body pairs; this covers ray hits.
' ============================================================

' --- HIT STATE ---
Public hit As Boolean

' --- IMPACT DATA ---
Public point  As Vector3
Public normal As Vector3
Public distance As Single

' --- OPTIONAL METADATA ---
Public entityID   As Long
Public MaterialID As Long
Public TriangleIndex As Long

' =========================
' INIT
' =========================
Private Sub Class_Initialize()
    hit = False
    Set point  = New Vector3
    Set normal = New Vector3
    distance = 0
    entityID = -1
    MaterialID = -1
    TriangleIndex = -1
End Sub

' =========================
' RESET (reuse without re-allocating)
' =========================
Public Sub Clear()
    hit = False
    point.SetValues 0, 0, 0
    normal.SetValues 0, 1, 0
    distance = 0
    entityID = -1
    MaterialID = -1
    TriangleIndex = -1
End Sub

' =========================
' COPY
' =========================
Public Sub CopyFrom(ByRef src As IntersectionResult)
    If src Is Nothing Then Exit Sub
    hit = src.hit
    point.CopyFrom src.point
    normal.CopyFrom src.normal
    distance = src.distance
    entityID = src.entityID
    MaterialID = src.MaterialID
    TriangleIndex = src.TriangleIndex
End Sub

' =========================
' CLEANUP
' =========================
Private Sub Class_Terminate()
    Set point  = Nothing
    Set normal = Nothing
End Sub

' =========================
' DEBUG
' =========================
Public Sub DebugPrint()
    Debug.Print "---- IntersectionResult ----"
    Debug.Print "Hit: "; hit
    If hit Then
        Debug.Print "Point:  "; point.x; point.y; point.z
        Debug.Print "Normal: "; normal.x; normal.y; normal.z
        Debug.Print "Dist:   "; distance
    End If
End Sub
