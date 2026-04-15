Option Explicit

' ============================================================
' Class       : Mesh
' Version     : v8.17
' Description : High-Level Mesh Wrapper (VBA ECS v7 Engine Core)
'               Synchronized with GL.bas v8.58 & MeshBuffer v8.1.1
' ============================================================

' =========================
' IDENTITY
' =========================

Public id As Long
Public name As String
Public Active As Boolean

' =========================
' CORE REFERENCES
' =========================

Public data As MeshData
Public buffer As MeshBuffer

' =========================
' BOUNDS (CACHE)
' =========================

Public BoundsMin As Vector3
Public BoundsMax As Vector3

' =========================
' INIT / TERMINATE
' =========================

Private Sub Class_Initialize()
    id = -1
    name = "New Mesh"
    Active = True

    Set data = Nothing
    Set buffer = Nothing

    Set BoundsMin = New Vector3
    Set BoundsMax = New Vector3
End Sub

Private Sub Class_Terminate()
    Dispose
End Sub

' =========================
' CREATE MESH
' =========================

Public Sub Create(ByRef MeshData As MeshData)
    Set data = MeshData
    If data Is Nothing Then Exit Sub

    Call SyncBounds
End Sub

' =========================
' ATTACH GPU BUFFER
' =========================

Public Sub BuildBuffer(Optional ByVal layout As VertexLayoutType = Layout_Position)
    If data Is Nothing Then Exit Sub
    
    Set buffer = New MeshBuffer
    Call buffer.LoadMesh(data, layout)
    Call buffer.Upload(data, layout)
End Sub

' =========================
' UPDATE MESH DATA
' =========================

Public Sub UpdateMesh(ByRef newData As MeshData)
    Set data = newData
    Call SyncBounds

    If Not buffer Is Nothing Then
        ' Synchronized with MeshBuffer.UpdateBuffer
        Call buffer.UpdateBuffer(data)
    End If
End Sub

' =========================
' SYNC BOUNDS FROM DATA
' =========================

Public Sub SyncBounds()
    If data Is Nothing Then Exit Sub

    Dim minV As Vector3
    Dim maxV As Vector3
    
    ' Vector3 objects must be initialized for GetBounds to fill them
    Set minV = New Vector3
    Set maxV = New Vector3

    Call data.GetBounds(minV, maxV)

    BoundsMin.SetValues minV.x, minV.y, minV.z
    BoundsMax.SetValues maxV.x, maxV.y, maxV.z
End Sub

' =========================
' DRAW MESH
' =========================

Public Sub Draw()
    If Not Active Then Exit Sub
    If data Is Nothing Then Exit Sub

    ' --------------------------------------------------------
    ' GPU PATH (Modern VBO/VAO)
    ' --------------------------------------------------------
    If Not buffer Is Nothing Then
        If buffer.IsUploaded Then
            buffer.Bind
            buffer.Draw
            Exit Sub
        End If
    End If

    ' --------------------------------------------------------
    ' CPU FALLBACK PATH (Legacy Immediate Mode)
    ' --------------------------------------------------------
    ' We pull arrays into local variables to avoid "Wrong number of arguments"
    ' errors caused by accessing Class Properties inside a loop.
    Dim i As Long
    Dim v() As Single: v = data.vertices
    Dim idxArr() As Long: idxArr = data.indices
    Dim vCount As Long: vCount = data.indexCount

    GL.glBegin GL_TRIANGLES

    For i = 0 To vCount - 1
        Dim idx As Long
        idx = idxArr(i)

        Dim vPos As Long
        vPos = idx * 3

        GL.glVertex3f v(vPos), v(vPos + 1), v(vPos + 2)
    Next i

    GL.glEnd
End Sub

' =========================
' RELEASE GPU
' =========================

Public Sub Dispose()
    If Not buffer Is Nothing Then
        Call buffer.Destroy
        Set buffer = Nothing
    End If
End Sub

' =========================
' HELPERS
' =========================

Public Sub GetBounds(ByRef minV As Vector3, ByRef maxV As Vector3)
    minV.SetValues BoundsMin.x, BoundsMin.y, BoundsMin.z
    maxV.SetValues BoundsMax.x, BoundsMax.y, BoundsMax.z
End Sub

Public Function IsValid() As Boolean
    If data Is Nothing Then Exit Function
    IsValid = data.IsValid
End Function

Public Sub Enable():  Active = True:  End Sub
Public Sub Disable(): Active = False: End Sub

' =========================
' DEBUG
' =========================

Public Sub DebugPrint()
    Debug.Print "---- MESH ----"
    Debug.Print "ID: "; id; " Name: "; name
    Debug.Print "Active: "; Active

    If Not data Is Nothing Then
        Debug.Print "Vertices: "; data.VertexCount
        Debug.Print "Indices: "; data.indexCount
    End If

    If Not buffer Is Nothing Then
        Debug.Print "GPU Uploaded: "; buffer.IsUploaded
    End If
End Sub


