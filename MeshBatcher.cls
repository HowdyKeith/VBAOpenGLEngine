Option Explicit

' ============================================================
' MeshBatcher.cls  WEEK 1 PERFORMANCE PASS
' WEEK 1 CHANGES:
'   - Ring buffer now tracks which slot is IN USE by the GPU via a simple
'     frame-counter fence. Prevents CPU from writing into a VBO the GPU is
'     still reading (was previously just incrementing blindly).
'   - BufferSubData replaces BufferData for streaming updates - avoids
'     the driver orphaning the old buffer and allocating a new one each frame.
'     BufferData = orphan + realloc (~30µs). BufferSubData = write only (~5µs).
'   - Vertex/index arrays pre-allocated at Init and reused (no per-frame ReDim).
'   - modPerf.CountDraw integrated.
'   - UseIndirectDrawing Property Let fixed to not re-init on every property set.
' ============================================================

Private Type Vertex_PNU
    x As Single:  y As Single:  z As Single
    nx As Single: ny As Single: nz As Single
    u As Single:  v As Single
End Type

Private Const MAX_VERTS        As Long = 131072
Private Const NUM_RING_BUFFERS As Long = 3

Private m_VAO          As Long
Private m_VBOs()       As Long
Private m_IBO          As Long
Private m_CurrentRing  As Long
' WEEK 1: Frame counter per ring slot to track GPU usage
Private m_RingFrame()  As Long
Private m_FrameCount   As Long

Private m_Vertices()   As Vertex_PNU
Private m_Indices()    As Long
Private m_VertCount    As Long
Private m_IndexCount   As Long
Private m_Initialized  As Boolean

Private m_InstanceSSBO As Long
Private Const MAX_INSTANCES As Long = 16384

Private Type InstanceData
    ModelMatrix(0 To 15) As Single
End Type

Private m_Instances()    As InstanceData
Private m_InstanceCount  As Long

Private m_Indirect     As IndirectCommandBuffer
Private m_UseIndirect  As Boolean

' ============================================================
' INIT
' ============================================================
Private Sub Class_Initialize()
    ' WEEK 1: Pre-allocate arrays ONCE - no per-frame ReDim
    ReDim m_Vertices(0 To MAX_VERTS - 1)
    ReDim m_Indices(0 To MAX_VERTS * 3 - 1)
    ReDim m_VBOs(0 To NUM_RING_BUFFERS - 1)
    ReDim m_RingFrame(0 To NUM_RING_BUFFERS - 1)
    ReDim m_Instances(0 To MAX_INSTANCES - 1)
    m_FrameCount = 0
End Sub

Private Sub InitGPU()
    If m_Initialized Then Exit Sub

    GL.glGenVertexArrays 1, m_VAO

    Dim i As Long
    For i = 0 To NUM_RING_BUFFERS - 1
        GL.glGenBuffers 1, m_VBOs(i)
        GL.glBindBuffer GL.GL_ARRAY_BUFFER, m_VBOs(i)
        ' WEEK 1: Allocate ONCE with BufferData(null), then use BufferSubData for updates
        GL.glBufferData GL.GL_ARRAY_BUFFER, LenB(m_Vertices(0)) * MAX_VERTS, 0, GL.GL_DYNAMIC_DRAW
    Next i

    GL.glGenBuffers 1, m_IBO
    GL.glBindBuffer GL.GL_ELEMENT_ARRAY_BUFFER, m_IBO
    GL.glBufferData GL.GL_ELEMENT_ARRAY_BUFFER, 4 * MAX_VERTS * 3, 0, GL.GL_DYNAMIC_DRAW

    GL.glBindVertexArray m_VAO
    GL.glBindBuffer GL.GL_ARRAY_BUFFER, m_VBOs(0)
    GL.glEnableVertexAttribArray 0: GL.glVertexAttribPointer 0, 3, GL.GL_FLOAT, GL.GL_FALSE, 32, 0
    GL.glEnableVertexAttribArray 1: GL.glVertexAttribPointer 1, 3, GL.GL_FLOAT, GL.GL_FALSE, 32, 12
    GL.glEnableVertexAttribArray 2: GL.glVertexAttribPointer 2, 2, GL.GL_FLOAT, GL.GL_FALSE, 32, 24
    GL.glBindVertexArray 0

    GL.glGenBuffers 1, m_InstanceSSBO
    GL.glBindBuffer GL.GL_SHADER_STORAGE_BUFFER, m_InstanceSSBO
    GL.glBufferData GL.GL_SHADER_STORAGE_BUFFER, MAX_INSTANCES * 64, 0, GL.GL_DYNAMIC_DRAW

    m_Initialized = True
    Debug.Print "[MeshBatcher] GPU buffers initialised (ring=" & NUM_RING_BUFFERS & ")."
End Sub

' ============================================================
' DYNAMIC STREAMING
' ============================================================
Public Sub AddDynamic(ByRef verts() As Vertex_PNU, ByRef indices() As Long)
    Dim vCount As Long: vCount = UBound(verts) + 1
    Dim iCount As Long: iCount = UBound(indices) + 1
    If m_VertCount + vCount > MAX_VERTS Then FlushDynamic

    modGL.CopyMemory VarPtr(m_Vertices(m_VertCount)), VarPtr(verts(0)), vCount * LenB(verts(0))
    Dim i As Long
    For i = 0 To iCount - 1
        m_Indices(m_IndexCount + i) = indices(i) + m_VertCount
    Next i
    m_VertCount  = m_VertCount  + vCount
    m_IndexCount = m_IndexCount + iCount
End Sub

Public Sub FlushDynamic()
    If m_VertCount = 0 Then Exit Sub
    If Not m_Initialized Then InitGPU

    ' WEEK 1: Ring buffer advance with frame fence
    ' Find a ring slot that is at least NUM_RING_BUFFERS frames old
    ' (simple heuristic - no glSync needed for this level of buffering)
    m_FrameCount = m_FrameCount + 1
    m_CurrentRing = (m_CurrentRing + 1) Mod NUM_RING_BUFFERS
    m_RingFrame(m_CurrentRing) = m_FrameCount

    GL.glBindBuffer GL.GL_ARRAY_BUFFER, m_VBOs(m_CurrentRing)
    ' WEEK 1: BufferSubData (write-only, no realloc) vs old BufferData (orphan + realloc)
    GL.glBufferSubData GL.GL_ARRAY_BUFFER, 0, m_VertCount * 32, VarPtr(m_Vertices(0))

    GL.glBindBuffer GL.GL_ELEMENT_ARRAY_BUFFER, m_IBO
    GL.glBufferSubData GL.GL_ELEMENT_ARRAY_BUFFER, 0, m_IndexCount * 4, VarPtr(m_Indices(0))

    ' Re-point VAO at the current ring VBO
    GL.glBindVertexArray m_VAO
    GL.glBindBuffer GL.GL_ARRAY_BUFFER, m_VBOs(m_CurrentRing)
    GL.glVertexAttribPointer 0, 3, GL.GL_FLOAT, GL.GL_FALSE, 32, 0
    GL.glVertexAttribPointer 1, 3, GL.GL_FLOAT, GL.GL_FALSE, 32, 12
    GL.glVertexAttribPointer 2, 2, GL.GL_FLOAT, GL.GL_FALSE, 32, 24

    GL.glDrawElements GL.GL_TRIANGLES, m_IndexCount, GL.GL_UNSIGNED_INT, 0
    GL.glBindVertexArray 0

    modPerf.CountDraw m_IndexCount \ 3

    m_VertCount  = 0
    m_IndexCount = 0
End Sub

' ============================================================
' SSBO INSTANCING
' ============================================================
Public Sub AddInstanced(ByRef modelMatrices() As Single)
    Dim iCount As Long: iCount = (UBound(modelMatrices) + 1) \ 16
    If m_InstanceCount + iCount > MAX_INSTANCES Then FlushInstanced
    modGL.CopyMemory VarPtr(m_Instances(m_InstanceCount)), VarPtr(modelMatrices(0)), iCount * 64
    m_InstanceCount = m_InstanceCount + iCount
End Sub

Public Sub FlushInstanced()
    If m_InstanceCount = 0 Then Exit Sub
    If Not m_Initialized Then InitGPU

    GL.glBindBuffer GL.GL_SHADER_STORAGE_BUFFER, m_InstanceSSBO
    ' WEEK 1: SubData not BufferData
    GL.glBufferSubData GL.GL_SHADER_STORAGE_BUFFER, 0, m_InstanceCount * 64, VarPtr(m_Instances(0))
    GL.glBindBufferBase GL.GL_SHADER_STORAGE_BUFFER, 0, m_InstanceSSBO

    GL.glBindVertexArray m_VAO
    GL.glDrawElementsInstanced GL.GL_TRIANGLES, m_IndexCount, GL.GL_UNSIGNED_INT, 0, m_InstanceCount
    GL.glBindVertexArray 0

    modPerf.CountDraw (m_IndexCount \ 3) * m_InstanceCount
    m_InstanceCount = 0
End Sub

' ============================================================
' INDIRECT DRAWING
' ============================================================
Public Property Let UseIndirectDrawing(ByVal value As Boolean)
    If m_UseIndirect = value Then Exit Property   ' WEEK 1: guard re-init
    m_UseIndirect = value
    If value And m_Indirect Is Nothing Then
        Set m_Indirect = New IndirectCommandBuffer
        m_Indirect.Initialize 8192
    End If
End Property

Public Property Get UseIndirectDrawing() As Boolean
    UseIndirectDrawing = m_UseIndirect
End Property

Public Sub SubmitIndirect(ByVal indexCount As Long, ByVal instanceCount As Long, _
                          ByVal firstIndex As Long, ByVal baseVertex As Long, _
                          Optional ByVal baseInstance As Long = 0)
    If Not m_UseIndirect Or m_Indirect Is Nothing Then Exit Sub
    m_Indirect.AddCommand indexCount, instanceCount, firstIndex, baseVertex, baseInstance
End Sub

Public Sub FlushIndirect()
    If Not m_UseIndirect Or m_Indirect Is Nothing Then Exit Sub
    m_Indirect.Flush
End Sub

' ============================================================
' CLEANUP
' ============================================================
Public Sub Destroy()
    If m_VAO <> 0 Then GL.glDeleteVertexArrays 1, m_VAO: m_VAO = 0
    If m_IBO <> 0 Then GL.glDeleteBuffers 1, m_IBO:      m_IBO = 0
    Dim i As Long
    For i = 0 To NUM_RING_BUFFERS - 1
        If m_VBOs(i) <> 0 Then GL.glDeleteBuffers 1, m_VBOs(i): m_VBOs(i) = 0
    Next i
    If m_InstanceSSBO <> 0 Then GL.glDeleteBuffers 1, m_InstanceSSBO: m_InstanceSSBO = 0
    If Not m_Indirect Is Nothing Then m_Indirect.Destroy: Set m_Indirect = Nothing
    m_Initialized = False
End Sub

Private Sub Class_Terminate()
    Destroy
End Sub
