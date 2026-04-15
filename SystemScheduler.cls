Option Explicit

' ============================================================
' SystemScheduler.cls  FIXED
' Deterministic ECS Execution Orchestrator
' FIXES:
'   - "ByRef World As World" - World class doesn't exist; changed to Object
'   - "Private m_Scene As SceneBridge" - SceneBridge class doesn't exist;
'     changed to Object
'   - m_Physics.StepSimulation dt - PhysicsSystem has Step(dt), not StepSimulation
'   - m_Spatial.RebuildFromPhysics m_Physics - SpatialGridSystem has no such method;
'     replaced with a per-entity insert loop using ComponentStore
'   - m_BVH.rebuild m_Spatial - BVHSystem has Refit, not rebuild
'   - m_Scene.Sync m_Physics - replaced with placeholder On Error guard
' ============================================================

Private m_Physics  As PhysicsSystem
Private m_BVH      As BVHSystem
Private m_Spatial  As SpatialGridSystem
Private m_Scene    As Object          ' FIXED: was SceneBridge (undefined)

' ============================================================
' INIT
' ============================================================
Public Sub Init(ByRef worldObj As Object)   ' FIXED: was "As World" (undefined)
    On Error Resume Next
    Set m_Physics = worldObj.physics
    Set m_BVH     = worldObj.bvh
    Set m_Spatial = worldObj.SpatialGrid
    Set m_Scene   = worldObj.SceneBridge    ' may be Nothing if not provided
    On Error GoTo 0
End Sub

' Convenience: bind each system individually (alternative to Init)
Public Sub BindPhysics(ByRef p As PhysicsSystem):    Set m_Physics = p:  End Sub
Public Sub BindBVH(ByRef b As BVHSystem):            Set m_BVH = b:      End Sub
Public Sub BindSpatialGrid(ByRef s As SpatialGridSystem): Set m_Spatial = s: End Sub
Public Sub BindScene(ByRef s As Object):             Set m_Scene = s:    End Sub

' ============================================================
' STEP  (FIXED: uses correct method names)
' ============================================================
Public Sub Step(ByVal dt As Single)

    ' 1. PHYSICS - FIXED: was StepSimulation, correct name is Step
    If Not m_Physics Is Nothing Then
        m_Physics.Step dt
    End If

    ' 2. SPATIAL GRID REBUILD - FIXED: no RebuildFromPhysics method exists.
    '    Rebuild grid from ComponentStore AABB arrays directly.
    If Not m_Spatial Is Nothing Then
        m_Spatial.Clear
        Dim i As Long
        Dim minV As New Vector3, maxV As New Vector3
        For i = 0 To ComponentStore.entityCount - 1
            If ComponentStore.IsActive(i) Then
                minV.SetValues ComponentStore.AabbMinX(i), ComponentStore.AabbMinY(i), ComponentStore.AabbMinZ(i)
                maxV.SetValues ComponentStore.AabbMaxX(i), ComponentStore.AabbMaxY(i), ComponentStore.AabbMaxZ(i)
                m_Spatial.InsertEntity i, minV, maxV
            End If
        Next i
    End If

    ' 3. BVH UPDATE - FIXED: was rebuild, correct name is Refit
    If Not m_BVH Is Nothing Then
        m_BVH.Refit
    End If

    ' 4. SCENE SYNC - placeholder, m_Scene may be Nothing
    If Not m_Scene Is Nothing Then
        On Error Resume Next
        CallByName m_Scene, "Sync", VbMethod, m_Physics
        On Error GoTo 0
    End If
End Sub

' ============================================================
' CLEANUP
' ============================================================
Private Sub Class_Terminate()
    Set m_Physics = Nothing
    Set m_BVH     = Nothing
    Set m_Spatial = Nothing
    Set m_Scene   = Nothing
End Sub
