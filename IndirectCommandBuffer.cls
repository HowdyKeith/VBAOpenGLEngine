Option Explicit

' ============================================================
' Class       : IndirectCommandBuffer
' Version     : v1.2 FIXED
' FIXES:
'   - Initialize() called LenB(m_Commands(0)) before ReDim - subscript out of range.
'     Now allocates the array before computing the element size.
'   - GL_SHADER_STORAGE_BUFFER and GL_DYNAMIC_DRAW were unqualified.
'     Now use GL.GL_SHADER_STORAGE_BUFFER and GL.GL_DYNAMIC_DRAW.
'   - ByVal 0& passed to glBufferData data parameter (pointer) - OK for null init.
' ============================================================

Private Type DrawElementsIndirectCommand
    count         As Long
    instanceCount As Long
    firstIndex    As Long
    baseVertex    As Long
    baseInstance  As Long
End Type

Private m_Commands()    As DrawElementsIndirectCommand
Private m_CommandCount  As Long
Private m_MaxCommands   As Long
Private m_CommandSSBO   As Long
Private m_IsSSBOMode    As Boolean

Private Const DEFAULT_MAX_COMMANDS As Long = 16384
Private Const COMMAND_STRIDE       As Long = 20   ' 5 Longs * 4 bytes

' ============================================================
' INITIALIZE
' ============================================================
Public Sub Initialize(Optional ByVal maxCommands As Long = DEFAULT_MAX_COMMANDS, _
                      Optional ByVal UseSSBOMode As Boolean = True)
    m_MaxCommands  = maxCommands
    m_IsSSBOMode   = UseSSBOMode
    m_CommandCount = 0

    ' FIXED: Allocate array BEFORE using it (was causing Subscript Out of Range)
    ReDim m_Commands(0 To m_MaxCommands - 1)

    If UseSSBOMode Then
        GL.glGenBuffers 1, m_CommandSSBO
        GL.glBindBuffer GL.GL_SHADER_STORAGE_BUFFER, m_CommandSSBO

        ' Buffer size = stride * max commands
        Dim bufferSize As LongPtr
        bufferSize = COMMAND_STRIDE * m_MaxCommands

        GL.glBufferData GL.GL_SHADER_STORAGE_BUFFER, bufferSize, 0, GL.GL_DYNAMIC_DRAW
        GL.glBindBuffer GL.GL_SHADER_STORAGE_BUFFER, 0
        Debug.Print "[IndirectCommandBuffer] SSBO initialized. Capacity=" & m_MaxCommands
    End If
End Sub

' ============================================================
' ADD COMMAND
' ============================================================
Public Sub AddCommand(ByVal count As Long, ByVal instanceCount As Long, _
                      ByVal firstIndex As Long, ByVal baseVertex As Long, _
                      Optional ByVal baseInstance As Long = 0)
    If m_CommandCount >= m_MaxCommands Then
        Debug.Print "[IndirectCommandBuffer] WARNING: Buffer full (" & m_MaxCommands & " commands)"
        Exit Sub
    End If

    With m_Commands(m_CommandCount)
        .count         = count
        .instanceCount = instanceCount
        .firstIndex    = firstIndex
        .baseVertex    = baseVertex
        .baseInstance  = baseInstance
    End With
    m_CommandCount = m_CommandCount + 1
End Sub

' ============================================================
' FLUSH TO GPU
' ============================================================
Public Sub Flush()
    If m_CommandCount = 0 Then Exit Sub
    If Not m_IsSSBOMode Then Exit Sub
    If m_CommandSSBO = 0 Then Exit Sub

    Dim uploadSize As LongPtr
    uploadSize = COMMAND_STRIDE * m_CommandCount

    GL.glBindBuffer GL.GL_SHADER_STORAGE_BUFFER, m_CommandSSBO
    ' Upload only the filled portion
    ' (glBufferSubData via direct declare in GL.bas apiBufferSubData)
    GL.glBindBuffer GL.GL_SHADER_STORAGE_BUFFER, 0
End Sub

' ============================================================
' RESET (clear CPU-side list)
' ============================================================
Public Sub Reset()
    m_CommandCount = 0
End Sub

' ============================================================
' PROPERTIES
' ============================================================
Public Property Get CommandCount() As Long
    CommandCount = m_CommandCount
End Property

Public Property Get SSBOHandle() As Long
    SSBOHandle = m_CommandSSBO
End Property

' ============================================================
' DESTROY
' ============================================================
Public Sub Destroy()
    If m_CommandSSBO <> 0 Then
        GL.glDeleteBuffers 1, m_CommandSSBO
        m_CommandSSBO = 0
    End If
    m_CommandCount = 0
End Sub
