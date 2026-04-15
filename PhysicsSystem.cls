Option Explicit

' =========================================================
' PhysicsSystem.v8.15.cls  FIXED
' BVH INTEGRATED CHARACTER PHYSICS
' FIXES:
'   - Missing CreateBody() method (called by Scene.cls)
'   - Missing GetBodyPosition() method (called by Scene.cls)
'   - RaycastDown now correctly typed with IntersectionResult (not CollisionResult)
'   - Step() referenced ComponentStore.IsActive/IsStatic as functions but they
'     are public arrays - fixed to array index access
' =========================================================

Public Gravity As Single
Private m_BVH As BVHSystem

' Simple body store (position only for now)
Private m_BodyCount As Long
Private m_BodyPosX() As Single
Private m_BodyPosY() As Single
Private m_BodyPosZ() As Single

Public Sub Init(ByRef bvh As BVHSystem)
    Set m_BVH = bvh
    Gravity = -9.81
    m_BodyCount = 0
End Sub

Private Sub Class_Initialize()
    Gravity = -9.81
    m_BodyCount = 0
End Sub

' =========================================================
' BODY MANAGEMENT (used by Scene.cls)
' =========================================================

Public Function CreateBody() As Long
    Dim i As Long
    i = m_BodyCount
    m_BodyCount = m_BodyCount + 1

    ReDim Preserve m_BodyPosX(0 To i)
    ReDim Preserve m_BodyPosY(0 To i)
    ReDim Preserve m_BodyPosZ(0 To i)

    m_BodyPosX(i) = 0
    m_BodyPosY(i) = 0
    m_BodyPosZ(i) = 0

    CreateBody = i
End Function

Public Function GetBodyPosition(ByVal bodyID As Long) As Vector3
    Dim v As New Vector3
    If bodyID >= 0 And bodyID < m_BodyCount Then
        v.SetValues m_BodyPosX(bodyID), m_BodyPosY(bodyID), m_BodyPosZ(bodyID)
    End If
    Set GetBodyPosition = v
End Function

Public Sub SetBodyPosition(ByVal bodyID As Long, ByVal x As Single, ByVal y As Single, ByVal z As Single)
    If bodyID >= 0 And bodyID < m_BodyCount Then
        m_BodyPosX(bodyID) = x
        m_BodyPosY(bodyID) = y
        m_BodyPosZ(bodyID) = z
    End If
End Sub

' =========================================================
' RAYCASTING  (FIXED: OutResult typed as IntersectionResult)
' =========================================================
Public Sub RaycastDown(ByRef Origin As Vector3, ByRef OutResult As IntersectionResult)
    OutResult.Clear

    ' 1. BVH raycast (if available)
    If Not m_BVH Is Nothing Then
        ' m_BVH.IntersectRay Origin, New Vector3(0,-1,0), OutResult  ' future hook
    End If

    ' 2. Fallback: infinite procedural floor at Y=0
    If Not OutResult.hit Then
        If Origin.y >= 0 Then
            OutResult.hit = True
            OutResult.point.SetValues Origin.x, 0, Origin.z
            OutResult.normal.SetValues 0, 1, 0
            OutResult.distance = Origin.y
        End If
    End If
End Sub

' =========================================================
' DYNAMIC UPDATE  (FIXED: array access instead of function calls)
' =========================================================
Public Sub Step(ByVal dt As Single)
    Dim i As Long, n As Long
    n = ComponentStore.entityCount

    For i = 0 To n - 1
        ' FIXED: IsActive and IsStatic are public arrays, not functions
        If ComponentStore.IsActive(i) Then
            If Not ComponentStore.IsStatic(i) Then
                ' Apply gravity to VelocityY
                ComponentStore.VelocityY(i) = ComponentStore.VelocityY(i) + Gravity * dt

                ' Integrate position
                ComponentStore.PositionX(i) = ComponentStore.PositionX(i) + ComponentStore.VelocityX(i) * dt
                ComponentStore.PositionY(i) = ComponentStore.PositionY(i) + ComponentStore.VelocityY(i) * dt
                ComponentStore.PositionZ(i) = ComponentStore.PositionZ(i) + ComponentStore.VelocityZ(i) * dt

                ' Hard floor
                If ComponentStore.PositionY(i) < 0 Then
                    ComponentStore.PositionY(i) = 0
                    ComponentStore.VelocityY(i) = 0
                End If
            End If
        End If
    Next i

    If Not m_BVH Is Nothing Then m_BVH.Refit
End Sub
