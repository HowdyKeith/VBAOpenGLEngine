Option Explicit

' ============================================================
' Scene.cls v1.4 FIXED
' FIXES: SystemMgr.Tick -> SYS_Tick; Matrix4 type; GetWorldPosition added
' ============================================================

Public SystemMgr  As Object
Public SceneGraph As Object
Public physics    As Object
Public Events     As Object
Public InputSys   As Object
Public Camera     As Object

Private Type EntityData
    id As Long:        name As String:   Active As Boolean
    TransformNode As Long: meshHandle As Long: SubMeshHandle As Long
    PhysicsBody As Long:   HasPhysics As Boolean: visible As Boolean
End Type

Private m_Entities() As EntityData
Private m_Count      As Long
Public IsRunning     As Boolean

Private Sub Class_Initialize()
    m_Count = 0: ReDim m_Entities(0 To 0): IsRunning = False
End Sub

Public Sub BindSystems(ByRef sysMgr As Object, ByRef graphRef As Object, _
    ByRef physicsSys As Object, ByRef eventSys As Object, _
    ByRef inputRef As Object, ByRef camRef As Object)
    Set SystemMgr=sysMgr: Set SceneGraph=graphRef: Set physics=physicsSys
    Set Events=eventSys:  Set InputSys=inputRef:   Set Camera=camRef
End Sub

Public Function CreateEntity(ByVal entityName As String) As Long
    m_Count = m_Count + 1
    ReDim Preserve m_Entities(0 To m_Count - 1)
    With m_Entities(m_Count - 1)
        .id=m_Count-1: .name=entityName: .Active=True
        .TransformNode=SceneGraph.CreateNode()
        .meshHandle=-1: .SubMeshHandle=-1: .PhysicsBody=-1
        .HasPhysics=False: .visible=True
    End With
    CreateEntity = m_Count - 1
End Function

Public Function GetEntityName(ByVal id As Long) As String
    If id<0 Or id>=m_Count Then Exit Function
    GetEntityName = m_Entities(id).name
End Function

Public Function IsActive(ByVal id As Long) As Boolean
    If id<0 Or id>=m_Count Then Exit Function
    IsActive = m_Entities(id).Active
End Function

Public Function entityCount() As Long: entityCount = m_Count: End Function

Public Sub SetPosition(ByVal id As Long, ByVal x As Single, ByVal y As Single, ByVal z As Single)
    If id<0 Or id>=m_Count Then Exit Sub
    SceneGraph.SetPosition m_Entities(id).TransformNode, x, y, z
    SceneGraph.MarkDirty m_Entities(id).TransformNode
End Sub

Public Sub SetRotation(ByVal id As Long, ByVal x As Single, ByVal y As Single, ByVal z As Single)
    If id<0 Or id>=m_Count Then Exit Sub
    SceneGraph.SetRotation m_Entities(id).TransformNode, x, y, z
    SceneGraph.MarkDirty m_Entities(id).TransformNode
End Sub

Public Sub SetScale(ByVal id As Long, ByVal x As Single, ByVal y As Single, ByVal z As Single)
    If id<0 Or id>=m_Count Then Exit Sub
    SceneGraph.SetScale m_Entities(id).TransformNode, x, y, z
    SceneGraph.MarkDirty m_Entities(id).TransformNode
End Sub

Public Function GetWorldPosition(ByVal id As Long) As Vector3
    Set GetWorldPosition = New Vector3
    If id<0 Or id>=m_Count Then Exit Function
    Set GetWorldPosition = SceneGraph.GetWorldPosition(m_Entities(id).TransformNode)
End Function

Public Sub AddPhysics(ByVal id As Long)
    If id<0 Or id>=m_Count Or physics Is Nothing Then Exit Sub
    Dim body As Long: body = physics.CreateBody()
    m_Entities(id).PhysicsBody=body: m_Entities(id).HasPhysics=True
End Sub

Private Sub SyncPhysicsToTransform()
    Dim i As Long
    For i = 0 To m_Count - 1
        If m_Entities(i).HasPhysics Then
            Dim pos As Vector3
            Set pos = physics.GetBodyPosition(m_Entities(i).PhysicsBody)
            SceneGraph.SetPosition m_Entities(i).TransformNode, pos.x, pos.y, pos.z
        End If
    Next i
End Sub

Public Sub Update(ByVal dt As Single)
    If Not IsRunning Then Exit Sub
    ' FIXED: SystemMgr.Tick -> SYS_Tick (SystemManager only has SYS_Tick)
    If Not SystemMgr Is Nothing Then
        On Error Resume Next: SystemMgr.SYS_Tick: On Error GoTo 0
    End If
    If Not SceneGraph Is Nothing Then SceneGraph.Update
    If Not physics Is Nothing Then SyncPhysicsToTransform
    If Not InputSys Is Nothing Then
        If InputSys.IsKeyPressed(vbKeyEscape) Then
            If Not Events Is Nothing Then Events.Emit "SCENE_EXIT", 0
        End If
    End If
End Sub

Public Sub Render()
    If Not IsRunning Then Exit Sub
    Dim i As Long
    For i = 0 To m_Count - 1
        If m_Entities(i).visible Then
            Dim worldMat As Matrix4
            worldMat = SceneGraph.GetWorldMatrix(m_Entities(i).TransformNode)
            If m_Entities(i).meshHandle <> -1 Then RenderEntityMesh i, worldMat
        End If
    Next i
End Sub

Private Sub RenderEntityMesh(ByVal id As Long, ByRef worldMat As Matrix4)
    ' Hook: MaterialSystem + ShaderProgram + MeshBatcher
End Sub

Public Sub SetVisible(ByVal id As Long, ByVal v As Boolean)
    If id<0 Or id>=m_Count Then Exit Sub: m_Entities(id).visible=v
End Sub

Public Sub DestroyEntity(ByVal id As Long)
    If id<0 Or id>=m_Count Then Exit Sub: m_Entities(id).Active=False
End Sub

Public Sub StartScene(): IsRunning=True:  End Sub
Public Sub StopScene():  IsRunning=False: End Sub

Public Sub Clear()
    m_Count=0: ReDim m_Entities(0 To 0): IsRunning=False
End Sub

Public Sub DebugPrint()
    Debug.Print "---- SCENE ---- Entities: "; m_Count; " Running: "; IsRunning
End Sub
