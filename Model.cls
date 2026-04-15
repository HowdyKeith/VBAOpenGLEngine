Option Explicit

' ============================================================
' Model.cls
' Renderable 3D Asset Container (VBA ECS v7 Engine Core)
' ============================================================

' =========================
' MODEL IDENTITY
' =========================

Public id As Long
Public name As String

Public Active As Boolean
Public visible As Boolean

' =========================
' MESH DATA
' =========================

Private m_SubMeshes() As Long
Private m_SubMeshCount As Long

' Optional per-submesh materials
Private m_Materials() As Long

' =========================
' TRANSFORM CACHE (OPTIONAL)
' =========================

Public Position As Vector3
Public Rotation As Vector3
'FIX
Public ScaleVec As Vector3

' =========================
' BOUNDING VOLUME (FOR BVH / CULLING)
' =========================

Public BoundsMin As Vector3
Public BoundsMax As Vector3

' =========================
' INITIALIZATION
' =========================

Private Sub Class_Initialize()

    id = -1
    name = ""

    Active = True
    visible = True

    m_SubMeshCount = 0
    ReDim m_SubMeshes(0 To 0)
    ReDim m_Materials(0 To 0)

    Set Position = New Vector3
    Set Rotation = New Vector3
    Set ScaleVec = New Vector3

    ScaleVec.x = 1
    ScaleVec.y = 1
    ScaleVec.z = 1

    Set BoundsMin = New Vector3
    Set BoundsMax = New Vector3

End Sub

' =========================
' INITIALIZE MODEL
' =========================

Public Sub Initialize(ByVal modelID As Long, ByVal modelName As String)

    id = modelID
    name = modelName

End Sub

' =========================
' ADD SUBMESH
' =========================

Public Sub AddSubMesh(ByVal meshID As Long, ByVal MaterialID As Long)

    m_SubMeshCount = m_SubMeshCount + 1

    ReDim Preserve m_SubMeshes(0 To m_SubMeshCount - 1)
    ReDim Preserve m_Materials(0 To m_SubMeshCount - 1)

    m_SubMeshes(m_SubMeshCount - 1) = meshID
    m_Materials(m_SubMeshCount - 1) = MaterialID

End Sub

' =========================
' GET SUBMESH COUNT
' =========================

Public Function SubMeshCount() As Long
    SubMeshCount = m_SubMeshCount
End Function

' =========================
' GET SUBMESH
' =========================

Public Function GetSubMesh(ByVal index As Long) As Long

    If index < 0 Or index >= m_SubMeshCount Then Exit Function

    GetSubMesh = m_SubMeshes(index)

End Function

' =========================
' GET MATERIAL
' =========================

Public Function GetMaterial(ByVal index As Long) As Long

    If index < 0 Or index >= m_SubMeshCount Then Exit Function

    GetMaterial = m_Materials(index)

End Function

' =========================
' SET TRANSFORM
' =========================

Public Sub SetPosition(ByVal x As Single, ByVal y As Single, ByVal z As Single)

    Position.x = x
    Position.y = y
    Position.z = z

End Sub

Public Sub SetRotation(ByVal x As Single, ByVal y As Single, ByVal z As Single)

    Rotation.x = x
    Rotation.y = y
    Rotation.z = z

End Sub

Public Sub SetScale(ByVal x As Single, ByVal y As Single, ByVal z As Single)

    ScaleVec.x = x
    ScaleVec.y = y
    ScaleVec.z = z

End Sub

' =========================
' UPDATE BOUNDS (FOR BVH / CULLING)
' =========================

Public Sub UpdateBounds(ByRef minV As Vector3, ByRef maxV As Vector3)

    BoundsMin.x = minV.x
    BoundsMin.y = minV.y
    BoundsMin.z = minV.z

    BoundsMax.x = maxV.x
    BoundsMax.y = maxV.y
    BoundsMax.z = maxV.z

End Sub

' =========================
' DRAW MODEL
' =========================

Public Sub Draw(ByRef renderer As Object)

    If Not Active Then Exit Sub
    If Not visible Then Exit Sub

    Dim i As Long

    For i = 0 To m_SubMeshCount - 1

        Dim meshID As Long
        Dim matID As Long

        meshID = m_SubMeshes(i)
        matID = m_Materials(i)

        ' ----------------------------------------------------
        ' Bind material (future MaterialSystem integration)
        ' ----------------------------------------------------

        Call renderer.BindMaterial(matID)

        ' ----------------------------------------------------
        ' Bind mesh
        ' ----------------------------------------------------

        Call renderer.DrawMesh(meshID, Position, Rotation, ScaleVec)

    Next i

End Sub

' =========================
' ENABLE / DISABLE
' =========================

Public Sub Enable()
    Active = True
End Sub

Public Sub Disable()
    Active = False
End Sub

' =========================
' SET VISIBILITY
' =========================

Public Sub SetVisible(ByVal value As Boolean)
    visible = value
End Sub

' =========================
' DEBUG
' =========================

Public Sub DebugPrint()

    Debug.Print "---- MODEL ----"
    Debug.Print "ID: "; id; " Name: "; name
    Debug.Print "SubMeshes: "; m_SubMeshCount
    Debug.Print "Active: "; Active; " Visible: "; visible
    Debug.Print "Bounds Min: "; BoundsMin.x; BoundsMin.y; BoundsMin.z
    Debug.Print "Bounds Max: "; BoundsMax.x; BoundsMax.y; BoundsMax.z

End Sub



