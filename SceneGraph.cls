Option Explicit

' =========================================================
' SceneGraph.cls  FIXED v15.1.0
' FIXES:
'   - "Matrix4" bare type now resolves via modTypes.bas (was undefined -> compile error)
'   - InstanceBuffer type now defined in modTypes.bas (was undefined -> compile error)
'   - GetWorldMatrix used "Empty" as both a variable name AND the VBA keyword Empty;
'     renamed variable to "emptyMat" to avoid the name collision
'   - LocalPosition/Rotation/Scale stored as Vector3 class objects but assigned with
'     "p = m_Nodes(n).LocalPosition" (value copy of an object) - fixed to use
'     module-level temp vars so BuildLocalMatrix reads correctly
'   - modMathUtils functions take GLMath.Mat4 but WorldMatrix is Matrix4;
'     VBA UDTs with identical layout are assignment-compatible so this works,
'     but we add explicit helpers via modTypes for clarity
' =========================================================

Private Type SceneNode
    id          As Long
    parent      As Long
    FirstChild  As Long
    NextSibling As Long
    LocalPosX   As Single
    LocalPosY   As Single
    LocalPosZ   As Single
    LocalRotX   As Single
    LocalRotY   As Single
    LocalRotZ   As Single
    LocalScaX   As Single
    LocalScaY   As Single
    LocalScaZ   As Single
    WorldMatrix As Matrix4
    meshID      As Long
    Renderable  As Boolean
    Dirty       As Boolean
End Type

Private m_Nodes() As SceneNode
Private m_Count   As Long

' =========================================================
' INIT
' =========================================================
Private Sub Class_Initialize()
    m_Count = 0
    ReDim m_Nodes(0 To 0)
End Sub

' =========================================================
' CREATE NODE
' =========================================================
Public Function CreateNode() As Long
    Dim i As Long
    i = m_Count
    m_Count = m_Count + 1
    ReDim Preserve m_Nodes(0 To m_Count - 1)

    With m_Nodes(i)
        .id          = i
        .parent      = -1
        .FirstChild  = -1
        .NextSibling = -1
        .LocalPosX = 0: .LocalPosY = 0: .LocalPosZ = 0
        .LocalRotX = 0: .LocalRotY = 0: .LocalRotZ = 0
        .LocalScaX = 1: .LocalScaY = 1: .LocalScaZ = 1
        .Renderable  = False
        .meshID      = -1
        .Dirty       = True
        IdentityMat4 .WorldMatrix
    End With

    CreateNode = i
End Function

' =========================================================
' TRANSFORM SETTERS
' =========================================================
Public Sub SetPosition(ByVal node As Long, ByVal x As Single, ByVal y As Single, ByVal z As Single)
    If node < 0 Or node >= m_Count Then Exit Sub
    m_Nodes(node).LocalPosX = x
    m_Nodes(node).LocalPosY = y
    m_Nodes(node).LocalPosZ = z
    m_Nodes(node).Dirty = True
End Sub

Public Sub SetRotation(ByVal node As Long, ByVal x As Single, ByVal y As Single, ByVal z As Single)
    If node < 0 Or node >= m_Count Then Exit Sub
    m_Nodes(node).LocalRotX = x
    m_Nodes(node).LocalRotY = y
    m_Nodes(node).LocalRotZ = z
    m_Nodes(node).Dirty = True
End Sub

Public Sub SetScale(ByVal node As Long, ByVal x As Single, ByVal y As Single, ByVal z As Single)
    If node < 0 Or node >= m_Count Then Exit Sub
    m_Nodes(node).LocalScaX = x
    m_Nodes(node).LocalScaY = y
    m_Nodes(node).LocalScaZ = z
    m_Nodes(node).Dirty = True
End Sub

Public Sub MarkDirty(ByVal node As Long)
    If node < 0 Or node >= m_Count Then Exit Sub
    m_Nodes(node).Dirty = True
End Sub

' =========================================================
' UPDATE (propagate dirty transforms down the hierarchy)
' =========================================================
Public Sub Update()
    Dim i As Long
    Dim identity As Matrix4
    IdentityMat4 identity

    For i = 0 To m_Count - 1
        If m_Nodes(i).parent = -1 Then
            UpdateNode i, identity
        End If
    Next i
End Sub

Private Sub UpdateNode(ByVal n As Long, ByRef parentMat As Matrix4)
    Dim c As Long
    If m_Nodes(n).Dirty Then
        BuildLocalMatrix n
        ' Multiply parent * local to get world
        Dim result As Matrix4
        MultiplyMat4 parentMat, m_Nodes(n).WorldMatrix, result
        m_Nodes(n).WorldMatrix = result
        m_Nodes(n).Dirty = False
    End If
    c = m_Nodes(n).FirstChild
    Do While c <> -1
        UpdateNode c, m_Nodes(n).WorldMatrix
        c = m_Nodes(c).NextSibling
    Loop
End Sub

Private Sub BuildLocalMatrix(ByVal n As Long)
    ' Build TRS (scale * rotate * translate) into WorldMatrix temporarily
    Dim mat As GLMath.Mat4
    modMathUtils.IdentityMatrix mat
    modMathUtils.ScaleMatrix    mat, m_Nodes(n).LocalScaX, m_Nodes(n).LocalScaY, m_Nodes(n).LocalScaZ
    modMathUtils.RotateMatrixX  mat, m_Nodes(n).LocalRotX
    modMathUtils.RotateMatrixY  mat, m_Nodes(n).LocalRotY
    modMathUtils.RotateMatrixZ  mat, m_Nodes(n).LocalRotZ
    modMathUtils.TranslateMatrix mat, m_Nodes(n).LocalPosX, m_Nodes(n).LocalPosY, m_Nodes(n).LocalPosZ
    ' Copy GLMath.Mat4 -> Matrix4
    modTypes.CopyMat4ToMatrix4 mat, m_Nodes(n).WorldMatrix
End Sub

' =========================================================
' GET WORLD MATRIX  (FIXED: renamed variable from "Empty" to "emptyMat")
' =========================================================
Public Function GetWorldMatrix(ByVal node As Long) As Matrix4
    Dim emptyMat As Matrix4
    IdentityMat4 emptyMat
    If node < 0 Or node >= m_Count Then
        GetWorldMatrix = emptyMat
        Exit Function
    End If
    GetWorldMatrix = m_Nodes(node).WorldMatrix
End Function

Public Function GetWorldPosition(ByVal node As Long) As Vector3
    Set GetWorldPosition = New Vector3
    If node < 0 Or node >= m_Count Then Exit Function
    GetWorldPosition.SetValues m_Nodes(node).WorldMatrix.m(12), _
                               m_Nodes(node).WorldMatrix.m(13), _
                               m_Nodes(node).WorldMatrix.m(14)
End Function

Public Function NodeCount() As Long
    NodeCount = m_Count
End Function

' =========================================================
' COLLECT RENDER INSTANCES  (FIXED: InstanceBuffer now defined in modTypes)
' =========================================================
Public Sub CollectRenderInstances(ByRef buffer As InstanceBuffer)
    Dim i As Long
    buffer.Count = 0
    If m_Count = 0 Then Exit Sub
    ReDim buffer.data(0 To m_Count - 1)

    For i = 0 To m_Count - 1
        If m_Nodes(i).Renderable Then
            buffer.data(buffer.Count).WorldMatrix = m_Nodes(i).WorldMatrix
            buffer.data(buffer.Count).meshID      = m_Nodes(i).meshID
            buffer.Count = buffer.Count + 1
        End If
    Next i

    If buffer.Count > 0 Then
        ReDim Preserve buffer.data(0 To buffer.Count - 1)
    End If
End Sub

' =========================================================
' PRIVATE MATH HELPERS (Matrix4 in-place - avoids GLMath.Mat4 conversions)
' =========================================================
Private Sub IdentityMat4(ByRef m As Matrix4)
    Dim i As Long
    For i = 0 To 15: m.m(i) = 0: Next i
    m.m(0) = 1: m.m(5) = 1: m.m(10) = 1: m.m(15) = 1
End Sub

Private Sub MultiplyMat4(ByRef a As Matrix4, ByRef b As Matrix4, ByRef result As Matrix4)
    Dim r As Long, c As Long, k As Long
    Dim tmp As Matrix4
    For r = 0 To 3
        For c = 0 To 3
            Dim s As Single: s = 0
            For k = 0 To 3
                s = s + a.m(r + k * 4) * b.m(k + c * 4)
            Next k
            tmp.m(r + c * 4) = s
        Next c
    Next r
    result = tmp
End Sub
