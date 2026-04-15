Option Explicit

' ============================================================
' Class      : RenderQueue
' Version    : v8.11
' Description: Collector for Culled Render Commands
' ============================================================

Private Type RenderCommand
    vbo As Long
    IBO As Long
    idxCount As Long
    matID As Long
    px As Single: py As Single: pz As Single
End Type

Private m_Queue() As RenderCommand
Private m_Count As Long

Public Sub Clear()
    m_Count = 0
    Erase m_Queue
End Sub

Public Sub Submit(ByVal vbo As Long, ByVal IBO As Long, ByVal idxCount As Long, _
                  ByVal matID As Long, ByVal px As Single, ByVal py As Single, ByVal pz As Single)
    
    ReDim Preserve m_Queue(m_Count)
    With m_Queue(m_Count)
        .vbo = vbo
        .IBO = IBO
        .idxCount = idxCount
        .matID = matID
        .px = px: .py = py: .pz = pz
    End With
    m_Count = m_Count + 1
End Sub

' Accessors for the Renderer
Public Function GetCount() As Long: GetCount = m_Count: End Function
' (Add further logic here for sorting by MaterialID to reduce state changes)




