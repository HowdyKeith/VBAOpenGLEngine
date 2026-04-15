Option Explicit

' ============================================================
' Class       : MeshData
' Version     : v8.6  (Layout-Aware)
' Description : CPU-Side Vertex and Index Storage + Layout Metadata
' ============================================================

Private m_Vertices() As Single
Private m_Indices() As Long
Private m_Layout As VertexLayoutType     ' ? NEW

Public Property Get vertices() As Single(): vertices = m_Vertices: End Property
Public Property Let vertices(ByRef v() As Single): m_Vertices = v: End Property

Public Property Get indices() As Long(): indices = m_Indices: End Property
Public Property Let indices(ByRef i() As Long): m_Indices = i: End Property

' ==================== NEW ====================
Public Property Get layout() As VertexLayoutType
    layout = m_Layout
End Property

Public Property Let layout(ByVal value As VertexLayoutType)
    m_Layout = value
End Property
' ============================================

Public Property Get VertexCount() As Long
    On Error Resume Next
    VertexCount = (UBound(m_Vertices) + 1) / GetStride(m_Layout)   ' ? updated to respect layout
    If Err.Number <> 0 Then VertexCount = 0
End Property

Public Property Get indexCount() As Long
    On Error Resume Next
    indexCount = UBound(m_Indices) + 1
    If Err.Number <> 0 Then indexCount = 0
End Property

Public Sub GetBounds(ByRef minV As Object, ByRef maxV As Object)
    ' ... (unchanged - works for any layout as long as position is first 3 floats)
    Dim vCount As Long: vCount = Me.VertexCount
    If vCount = 0 Then Exit Sub
    Dim i As Long, vx As Single, vy As Single, vz As Single
    minV.x = m_Vertices(0): maxV.x = m_Vertices(0)
    minV.y = m_Vertices(1): maxV.y = m_Vertices(1)
    minV.z = m_Vertices(2): maxV.z = m_Vertices(2)
    For i = GetStride(m_Layout) To (vCount * GetStride(m_Layout)) - 1 Step GetStride(m_Layout)
        vx = m_Vertices(i): vy = m_Vertices(i + 1): vz = m_Vertices(i + 2)
        If vx < minV.x Then minV.x = vx Else If vx > maxV.x Then maxV.x = vx
        If vy < minV.y Then minV.y = vy Else If vy > maxV.y Then maxV.y = vy
        If vz < minV.z Then minV.z = vz Else If vz > maxV.z Then maxV.z = vz
    Next i
End Sub

Public Function IsValid() As Boolean
    On Error Resume Next
    IsValid = (UBound(m_Vertices) >= 0 And UBound(m_Indices) >= 0)
End Function

' Optional helper
Public Sub SetDefaultLayout()
    If m_Layout = 0 Then m_Layout = Layout_PositionNormal   ' sensible default
End Sub
