Option Explicit
' =========================================================
' Class: TransformComponent
' Version: v8.1.0 (FIXED - SyncToScene added)
' =========================================================
Public Position    As Vector3
Public Rotation    As Vector3
Public ScaleVec    As Vector3
Public SceneNodeID As Long

Private Sub Class_Initialize()
    Set Position = New Vector3
    Set Rotation = New Vector3
    Set ScaleVec = New Vector3
    ScaleVec.x = 1
    ScaleVec.y = 1
    ScaleVec.z = 1
    SceneNodeID = -1
End Sub

' =========================================================
' SYNC TO SCENE GRAPH (called by TransformSystem each frame)
' =========================================================
Public Sub SyncToScene(ByRef graph As SceneGraph)
    If SceneNodeID < 0 Then Exit Sub
    If graph Is Nothing Then Exit Sub
    graph.SetPosition SceneNodeID, Position.x, Position.y, Position.z
    graph.SetRotation SceneNodeID, Rotation.x, Rotation.y, Rotation.z
    graph.SetScale SceneNodeID, ScaleVec.x, ScaleVec.y, ScaleVec.z
    graph.MarkDirty SceneNodeID
End Sub

' =========================================================
' HELPERS
' =========================================================
Public Sub SetPosition(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Position.SetValues x, y, z
End Sub

Public Sub SetRotation(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Rotation.SetValues x, y, z
End Sub

Public Sub SetScale(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    ScaleVec.SetValues x, y, z
End Sub

Private Sub Class_Terminate()
    Set Position = Nothing
    Set Rotation = Nothing
    Set ScaleVec = Nothing
End Sub

