Option Explicit

' ============================================================
' CollisionPipeline.v8.1.bas  FIXED
' ECS v8 Full Collision Orchestrator
' FIXES:
'   - "Public grid As SpatialGrid" - class is named SpatialGridSystem, not SpatialGrid
'   - "World.count" - PhysicsSystem has no .count property;
'     entity count comes from ComponentStore.entityCount
'   - "World.AabbMinX(i)" etc - those arrays live in ComponentStore, not PhysicsSystem
'   - "World.IsStatic(a)" - IsStatic is a ComponentStore array, not a method on PhysicsSystem
'   - "World.VelocityY(a)" - VelocityY arrays are in ComponentStore
'   - grid.Insert called with 2D args but SpatialGridSystem.InsertEntity takes Vector3 objects
' ============================================================

Public World As PhysicsSystem
Public grid  As SpatialGridSystem    ' FIXED: was SpatialGrid (wrong name)
Public bvh   As BVHSystem

' ============================================================
' MAIN PIPELINE STEP
' ============================================================
Public Sub StepCollision()
    Dim i As Long
    Dim n As Long
    n = ComponentStore.entityCount  ' FIXED: was World.count

    ' =====================================================
    ' 1. BROADPHASE (SPATIAL GRID)
    ' =====================================================
    If Not grid Is Nothing Then
        grid.Clear

        Dim minV As New Vector3
        Dim maxV As New Vector3

        For i = 0 To n - 1
            If ComponentStore.IsActive(i) Then
                ' FIXED: use ComponentStore arrays, not World.AabbMinX etc.
                minV.SetValues ComponentStore.AabbMinX(i), ComponentStore.AabbMinY(i), ComponentStore.AabbMinZ(i)
                maxV.SetValues ComponentStore.AabbMaxX(i), ComponentStore.AabbMaxY(i), ComponentStore.AabbMaxZ(i)
                grid.InsertEntity i, minV, maxV
            End If
        Next i
    End If

    ' =====================================================
    ' 2. MIDPHASE (BVH REFIT)
    ' =====================================================
    If Not bvh Is Nothing Then bvh.Refit

    ' =====================================================
    ' 3. NARROW PHASE (AABB PAIR TESTING)
    ' =====================================================
    Dim a As Long, b As Long
    For a = 0 To n - 1
        ' FIXED: ComponentStore.IsActive(a) is an array element, not a function
        If ComponentStore.IsActive(a) Then
            For b = a + 1 To n - 1
                If ComponentStore.IsActive(b) Then
                    ' Broad AABB overlap test
                    If Not (ComponentStore.AabbMaxX(a) < ComponentStore.AabbMinX(b) Or _
                            ComponentStore.AabbMinX(a) > ComponentStore.AabbMaxX(b) Or _
                            ComponentStore.AabbMaxY(a) < ComponentStore.AabbMinY(b) Or _
                            ComponentStore.AabbMinY(a) > ComponentStore.AabbMaxY(b) Or _
                            ComponentStore.AabbMaxZ(a) < ComponentStore.AabbMinZ(b) Or _
                            ComponentStore.AabbMinZ(a) > ComponentStore.AabbMaxZ(b)) Then
                        ResolvePair a, b
                    End If
                End If
            Next b
        End If
    Next a
End Sub

' ============================================================
' RESOLUTION HOOK
' FIXED: World.IsStatic / World.VelocityY -> ComponentStore arrays
' ============================================================
Private Sub ResolvePair(ByVal a As Long, ByVal b As Long)
    ' FIXED: ComponentStore.IsStatic is an array
    If ComponentStore.IsStatic(a) And ComponentStore.IsStatic(b) Then Exit Sub

    ' Simple elastic bounce response
    If Not ComponentStore.IsStatic(a) Then
        ComponentStore.VelocityY(a) = -ComponentStore.VelocityY(a) * 0.5
    End If
    If Not ComponentStore.IsStatic(b) Then
        ComponentStore.VelocityY(b) = -ComponentStore.VelocityY(b) * 0.5
    End If
End Sub
