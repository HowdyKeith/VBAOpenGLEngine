' ComponentManager.cls (v1.0 ECS HARDENED)
Option Explicit

Private m_Transforms As Collection

Private Sub Class_Initialize()
    Set m_Transforms = New Collection
End Sub

' =========================================================
' ECS COMPONENT API
' =========================================================

Public Sub ECS_AddTransform(ByVal entityID As Long, ByVal t As TransformComponent)
    m_Transforms.Add t, CStr(entityID)
End Sub

Public Function ECS_GetTransform(ByVal entityID As Long) As TransformComponent

    On Error Resume Next
    Set ECS_GetTransform = m_Transforms(CStr(entityID))
    On Error GoTo 0

End Function

Public Sub ECS_RemoveTransform(ByVal entityID As Long)

    On Error Resume Next
    m_Transforms.Remove CStr(entityID)
    On Error GoTo 0

End Sub

Public Function ECS_HasTransform(ByVal entityID As Long) As Boolean

    On Error Resume Next
    ECS_HasTransform = Not m_Transforms(CStr(entityID)) Is Nothing
    On Error GoTo 0

End Function


