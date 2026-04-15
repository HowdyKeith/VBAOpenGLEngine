' EntityManager.cls (v1.0 ECS HARDENED)
Option Explicit

Private m_Active() As Boolean
Private m_Generation() As Long
Private m_FreeList() As Long

Private m_Count As Long
Private m_Capacity As Long
Private m_FreeCount As Long

Private Sub Class_Initialize()

    m_Capacity = 1024

    ReDim m_Active(0 To m_Capacity - 1)
    ReDim m_Generation(0 To m_Capacity - 1)
    ReDim m_FreeList(0 To m_Capacity - 1)

End Sub

' =========================================================
' ECS LIFECYCLE API (HARDENED)
' =========================================================

Public Function ECS_CreateEntity() As Long

    Dim id As Long

    If m_FreeCount > 0 Then
        m_FreeCount = m_FreeCount - 1
        id = m_FreeList(m_FreeCount)
    Else
        If m_Count >= m_Capacity Then ExpandCapacity
        id = m_Count
        m_Count = m_Count + 1
    End If

    m_Active(id) = True
    m_Generation(id) = m_Generation(id) + 1

    ECS_CreateEntity = id

End Function

Public Sub ECS_DestroyEntity(ByVal entityID As Long)

    If entityID < 0 Or entityID >= m_Count Then Exit Sub

    m_Active(entityID) = False

    m_FreeList(m_FreeCount) = entityID
    m_FreeCount = m_FreeCount + 1

End Sub

Public Function ECS_IsValid(ByVal entityID As Long) As Boolean

    If entityID < 0 Or entityID >= m_Count Then Exit Function
    ECS_IsValid = m_Active(entityID)

End Function

Private Sub ExpandCapacity()

    Dim newCap As Long
    newCap = m_Capacity * 2

    ReDim Preserve m_Active(0 To newCap - 1)
    ReDim Preserve m_Generation(0 To newCap - 1)
    ReDim Preserve m_FreeList(0 To newCap - 1)

    m_Capacity = newCap

End Sub


