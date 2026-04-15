Option Explicit

' ============================================================
' Raycast.cls  FIXED
' Scene Query + Physics Raycasting
' FIXES:
'   - LastHit was "As RayHit" where RayHit tried to hold Vector3 object
'     references inside a .bas UDT - illegal in VBA. Replaced with
'     RayHitFlat (flat Singles) + separate HitPoint/HitNormal Vector3 fields
'   - Scene.GetAABB(i) doesn't exist on Scene - replaced with ComponentStore
'     AABB array access which is the correct ECS pattern
'   - Scene.GetWorldPosition(i) doesn't exist - replaced with SceneGraph
'     GetWorldPosition via the scene's SceneGraph reference
'   - Swap() conflicted with VBA internal - renamed SwapSingle()
' ============================================================

Public Origin      As Vector3
Public direction   As Vector3
Public maxDistance As Single

' FIXED: flat struct for hit result (no object refs in UDT)
Public LastHit     As RayHitFlat
' Separate Vector3 objects for hit point and normal
Public HitPoint    As Vector3
Public HitNormal   As Vector3

Private Sub Class_Initialize()
    Set Origin    = New Vector3
    Set direction = New Vector3
    Set HitPoint  = New Vector3
    Set HitNormal = New Vector3
    maxDistance       = 1000#
    LastHit.hit       = False
    LastHit.distance  = 0
    LastHit.entityID  = -1
End Sub

Private Sub Class_Terminate()
    Set Origin    = Nothing
    Set direction = Nothing
    Set HitPoint  = Nothing
    Set HitNormal = Nothing
End Sub

' ============================================================
' SET RAY
' ============================================================
Public Sub SetRay(ByVal ox As Single, ByVal oy As Single, ByVal oz As Single, _
                  ByVal dx As Single, ByVal dy As Single, ByVal dz As Single)
    Origin.SetValues ox, oy, oz
    direction.SetValues dx, dy, dz
    NormalizeDir direction
End Sub

' ============================================================
' MAIN CAST - tests against all active ECS entities
' Uses ComponentStore AABB data (ECS pattern, not Scene.GetAABB)
' ============================================================
Public Function Cast(ByRef sceneRef As Scene) As Boolean
    Dim i As Long
    Dim bestDist As Single: bestDist = maxDistance

    LastHit.hit      = False
    LastHit.distance = maxDistance
    LastHit.entityID = -1

    Dim n As Long: n = ComponentStore.entityCount
    For i = 0 To n - 1
        If ComponentStore.IsActive(i) Then
            ' FIXED: use ComponentStore AABB arrays directly
            If RayAABBTest(ComponentStore.AabbMinX(i), ComponentStore.AabbMinY(i), ComponentStore.AabbMinZ(i), _
                           ComponentStore.AabbMaxX(i), ComponentStore.AabbMaxY(i), ComponentStore.AabbMaxZ(i)) Then

                Dim dist As Single
                dist = ComputeEntityDist(i)
                If dist < bestDist Then
                    bestDist         = dist
                    LastHit.hit      = True
                    LastHit.distance = dist
                    LastHit.entityID = i
                    ' Compute hit point
                    LastHit.px = Origin.x + direction.x * dist
                    LastHit.py = Origin.y + direction.y * dist
                    LastHit.pz = Origin.z + direction.z * dist
                    ' Copy to public Vector3 fields for convenience
                    HitPoint.SetValues LastHit.px, LastHit.py, LastHit.pz
                End If
            End If
        End If
    Next i
    Cast = LastHit.hit
End Function

' ============================================================
' AABB SLAB TEST
' ============================================================
Private Function RayAABBTest(minX As Single, minY As Single, minZ As Single, _
                              maxX As Single, maxY As Single, maxZ As Single) As Boolean
    Dim t1 As Single, t2 As Single
    Dim tmin As Single: tmin = -1E+20
    Dim tmax As Single: tmax = 1E+20

    If direction.x <> 0 Then
        t1 = (minX - Origin.x) / direction.x
        t2 = (maxX - Origin.x) / direction.x
        If t1 > t2 Then SwapSingle t1, t2
        If t1 > tmin Then tmin = t1
        If t2 < tmax Then tmax = t2
    End If
    If direction.y <> 0 Then
        t1 = (minY - Origin.y) / direction.y
        t2 = (maxY - Origin.y) / direction.y
        If t1 > t2 Then SwapSingle t1, t2
        If t1 > tmin Then tmin = t1
        If t2 < tmax Then tmax = t2
    End If
    If direction.z <> 0 Then
        t1 = (minZ - Origin.z) / direction.z
        t2 = (maxZ - Origin.z) / direction.z
        If t1 > t2 Then SwapSingle t1, t2
        If t1 > tmin Then tmin = t1
        If t2 < tmax Then tmax = t2
    End If
    RayAABBTest = (tmax >= tmin And tmax > 0)
End Function

' ============================================================
' DISTANCE TO ENTITY CENTRE (from ComponentStore position)
' ============================================================
Private Function ComputeEntityDist(ByVal entityID As Long) As Single
    Dim dx As Single, dy As Single, dz As Single
    dx = ComponentStore.PositionX(entityID) - Origin.x
    dy = ComponentStore.PositionY(entityID) - Origin.y
    dz = ComponentStore.PositionZ(entityID) - Origin.z
    ComputeEntityDist = Sqr(dx * dx + dy * dy + dz * dz)
End Function

' ============================================================
' HELPERS
' ============================================================
Private Sub NormalizeDir(ByRef v As Vector3)
    Dim len As Single: len = Sqr(v.x * v.x + v.y * v.y + v.z * v.z)
    If len > 0 Then v.x = v.x / len: v.y = v.y / len: v.z = v.z / len
End Sub

' FIXED: renamed from Swap to SwapSingle (Swap is a VBA keyword in some hosts)
Private Sub SwapSingle(ByRef a As Single, ByRef b As Single)
    Dim tmp As Single: tmp = a: a = b: b = tmp
End Sub

' ============================================================
' DEBUG
' ============================================================
Public Sub DebugPrint()
    Debug.Print "---- RAYCAST ----"
    Debug.Print "Origin:    "; Origin.x; Origin.y; Origin.z
    Debug.Print "Direction: "; direction.x; direction.y; direction.z
    Debug.Print "Hit: "; LastHit.hit
    If LastHit.hit Then
        Debug.Print "  EntityID: "; LastHit.entityID
        Debug.Print "  Distance: "; LastHit.distance
        Debug.Print "  Point:    "; LastHit.px; LastHit.py; LastHit.pz
    End If
End Sub
