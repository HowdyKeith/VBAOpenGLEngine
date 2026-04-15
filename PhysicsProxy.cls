Option Explicit
' ============================================================
' PhysicsProxy.cls
' ECS <-> Physics Bridge Layer (v8 FIXED)
' RigidBody removed - delegates to ComponentStore SoA arrays
' and PhysicsSystem for simulation stepping.
' ============================================================

' =========================
' REFERENCES
' =========================
Public physics As PhysicsSystem
Public SceneRef As Scene

' =========================
' PROXY STATE
' =========================
Private m_EntityID As Long
Private m_Active   As Boolean

' =========================
' INIT
' =========================
Private Sub Class_Initialize()
    m_EntityID = -1
    m_Active = False
    Set physics = Nothing
    Set SceneRef = Nothing
End Sub

' =========================
' ATTACH TO ENTITY
' =========================
Public Sub Attach(ByVal entityID As Long, ByRef physicsWorld As PhysicsSystem)
    m_EntityID = entityID
    Set physics = physicsWorld
    m_Active = (m_EntityID >= 0 And Not physics Is Nothing)
End Sub

' =========================
' DETACH
' =========================
Public Sub Detach()
    m_EntityID = -1
    m_Active = False
    Set physics = Nothing
End Sub

' =========================
' APPLY FORCE (accumulates into velocity this frame)
' =========================
Public Sub AddForce(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    If Not m_Active Then Exit Sub
    ComponentStore.VelocityX(m_EntityID) = ComponentStore.VelocityX(m_EntityID) + x
    ComponentStore.VelocityY(m_EntityID) = ComponentStore.VelocityY(m_EntityID) + y
    ComponentStore.VelocityZ(m_EntityID) = ComponentStore.VelocityZ(m_EntityID) + z
End Sub

' =========================
' APPLY IMPULSE (same as AddForce in this simple integrator)
' =========================
Public Sub AddImpulse(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    AddForce x, y, z
End Sub

' =========================
' SET VELOCITY
' =========================
Public Sub SetVelocity(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    If Not m_Active Then Exit Sub
    ComponentStore.VelocityX(m_EntityID) = x
    ComponentStore.VelocityY(m_EntityID) = y
    ComponentStore.VelocityZ(m_EntityID) = z
End Sub

' =========================
' GET VELOCITY
' =========================
Public Function GetVelocity() As Vector3
    Set GetVelocity = New Vector3
    If Not m_Active Then Exit Function
    GetVelocity.SetValues ComponentStore.VelocityX(m_EntityID), _
                          ComponentStore.VelocityY(m_EntityID), _
                          ComponentStore.VelocityZ(m_EntityID)
End Function

' =========================
' GET POSITION
' =========================
Public Function GetPosition() As Vector3
    Set GetPosition = New Vector3
    If Not m_Active Then Exit Function
    GetPosition.SetValues ComponentStore.PositionX(m_EntityID), _
                          ComponentStore.PositionY(m_EntityID), _
                          ComponentStore.PositionZ(m_EntityID)
End Function

' =========================
' TELEPORT
' =========================
Public Sub SetPosition(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    If Not m_Active Then Exit Sub
    ComponentStore.PositionX(m_EntityID) = x
    ComponentStore.PositionY(m_EntityID) = y
    ComponentStore.PositionZ(m_EntityID) = z
    ComponentStore.VelocityX(m_EntityID) = 0
    ComponentStore.VelocityY(m_EntityID) = 0
    ComponentStore.VelocityZ(m_EntityID) = 0
End Sub

' =========================
' SYNC CACHE (called after PhysicsSystem.Step)
' =========================
Public Sub Sync()
    ' No-op: ComponentStore IS the authoritative state.
    ' This exists for API compatibility with callers expecting a Sync call.
End Sub

' =========================
' ENABLE / DISABLE
' =========================
Public Sub Enable():  m_Active = (m_EntityID >= 0): End Sub
Public Sub Disable(): m_Active = False: End Sub

' =========================
' STATUS
' =========================
Public Function IsActive() As Boolean:  IsActive = m_Active:     End Function
Public Function BodyID() As Long:       BodyID = m_EntityID:     End Function
Public Function entityID() As Long:     entityID = m_EntityID:   End Function

' =========================
' DEBUG
' =========================
Public Sub DebugPrint()
    Debug.Print "---- PHYSICS PROXY ----"
    Debug.Print "Entity: "; m_EntityID
    Debug.Print "Active: "; m_Active
    If Not m_Active Then Exit Sub
    Debug.Print "Pos: "; ComponentStore.PositionX(m_EntityID); _
                        ComponentStore.PositionY(m_EntityID); _
                        ComponentStore.PositionZ(m_EntityID)
    Debug.Print "Vel: "; ComponentStore.VelocityX(m_EntityID); _
                        ComponentStore.VelocityY(m_EntityID); _
                        ComponentStore.VelocityZ(m_EntityID)
End Sub

' =========================
' CLEANUP
' =========================
Private Sub Class_Terminate()
    Set physics = Nothing
    Set SceneRef = Nothing
End Sub


