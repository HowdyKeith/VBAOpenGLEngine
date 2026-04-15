Option Explicit

' ============================================================
' SpatialQuerySystem.cls (ECS v8 PIPELINE)
' ============================================================

Public grid As SpatialGridSystem
Public bvh As BVHSystem

' =========================
' RAY QUERY PIPELINE
' =========================

Public Function RayQuery(ByRef Origin As Vector3, _
                         ByRef dir As Vector3, _
                         ByRef colliders As Collection) As Collection

    Dim results As Collection
    Set results = New Collection

    Dim candidateIDs As Collection

    ' -------------------------
    ' STEP 1: BROADPHASE GRID
    ' -------------------------
    Dim minV As Vector3, maxV As Vector3

    Set minV = New Vector3
    Set maxV = New Vector3

    minV.x = Origin.x - 10
    minV.y = Origin.y - 10
    minV.z = Origin.z - 10

    maxV.x = Origin.x + 10
    maxV.y = Origin.y + 10
    maxV.z = Origin.z + 10

    grid.QueryAABB minV, maxV, candidateIDs

    ' -------------------------
    ' STEP 2: BVH + TRI TEST
    ' -------------------------
    Dim i As Long
    Dim id As Variant

    For Each id In candidateIDs

        Dim collider As MeshCollider
        Set collider = colliders(id)

        Dim dist As Single

        If collider.RayIntersect(Origin, dir, dist) Then
            results.Add id
        End If

    Next id

    Set RayQuery = results

End Function




