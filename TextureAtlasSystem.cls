Option Explicit

' ============================================================
' TextureAtlasSystem.v8.cls
' GPU Texture Atlas Manager (ECS v8 Clean Version)
' ============================================================
' - No UDTs (VBA class-safe)
' - Object-based region storage
' - GPU batch ready
' - Safe ECS integration
' ============================================================

' =========================
' ATLAS SETTINGS
' =========================

Private m_AtlasWidth As Long
Private m_AtlasHeight As Long

Private m_CurrentX As Long
Private m_CurrentY As Long
Private m_RowHeight As Long

' =========================
' GPU TEXTURE
' =========================

Public AtlasTexture As Texture2D

' =========================
' REGION STORAGE (ECS SAFE)
' =========================

Private m_Regions As Collection
' key -> AtlasRegion (class object)

' ============================================================
' ATLAS REGION CLASS (must exist as AtlasRegion.cls)
' ============================================================
' Public x As Single
' Public y As Single
' Public W As Single
' Public H As Single

' =========================
' INIT
' =========================

Private Sub Class_Initialize()

    Set m_Regions = New Collection

    m_AtlasWidth = 1024
    m_AtlasHeight = 1024

    m_CurrentX = 0
    m_CurrentY = 0
    m_RowHeight = 0

    Set AtlasTexture = New Texture2D
    AtlasTexture.CreateEmpty m_AtlasWidth, m_AtlasHeight

End Sub

' ============================================================
' ADD TEXTURE TO ATLAS
' ============================================================

Public Function AddTexture( _
    ByVal key As String, _
    ByVal tex As Texture2D) As Boolean

    Dim region As AtlasRegion

    If tex Is Nothing Then Exit Function

    ' -------------------------
    ' ROW PACKING LOGIC
    ' -------------------------

    If tex.width + m_CurrentX > m_AtlasWidth Then
        m_CurrentX = 0
        m_CurrentY = m_CurrentY + m_RowHeight
        m_RowHeight = 0
    End If

    If m_CurrentY + tex.Height > m_AtlasHeight Then
        AddTexture = False
        Exit Function
    End If

    ' -------------------------
    ' CREATE REGION OBJECT
    ' -------------------------

    Set region = New AtlasRegion

    region.x = m_CurrentX / m_AtlasWidth
    region.y = m_CurrentY / m_AtlasHeight
    region.w = tex.width / m_AtlasWidth
    region.h = tex.Height / m_AtlasHeight

    ' -------------------------
    ' STORE REGION (KEYED)
    ' -------------------------

    On Error Resume Next
    m_Regions.Add region, key
    On Error GoTo 0

    ' -------------------------
    ' BLIT INTO ATLAS BUFFER
    ' -------------------------

    BlitToAtlas tex, m_CurrentX, m_CurrentY

    ' advance cursor
    m_CurrentX = m_CurrentX + tex.width

    If tex.Height > m_RowHeight Then
        m_RowHeight = tex.Height
    End If

    AddTexture = True

End Function

' ============================================================
' GET UV REGION (SAFE OBJECT RETURN)
' ============================================================

Public Function GetRegion(ByVal key As String) As AtlasRegion

    On Error GoTo fail

    Set GetRegion = m_Regions(key)
    Exit Function

fail:
    Set GetRegion = Nothing

End Function

' ============================================================
' BIND ATLAS
' ============================================================

Public Sub Bind(Optional ByVal slot As Long = 0)

    If Not AtlasTexture Is Nothing Then
        AtlasTexture.Bind slot
    End If

End Sub

' ============================================================
' BLIT INTO ATLAS (CPU SIDE COPY)
' ============================================================

Private Sub BlitToAtlas(ByVal tex As Texture2D, ByVal x As Long, ByVal y As Long)

    Dim dx As Long, dy As Long
    Dim srcIndex As Long, dstIndex As Long

    If tex Is Nothing Then Exit Sub
    If AtlasTexture Is Nothing Then Exit Sub

    For dy = 0 To tex.Height - 1
        For dx = 0 To tex.width - 1

            srcIndex = (dy * tex.width + dx) * 4
            dstIndex = ((y + dy) * m_AtlasWidth + (x + dx)) * 4

            ' RGBA copy (assumes m_Pixels is exposed internally)
            AtlasTexture.m_Pixels(dstIndex) = tex.m_Pixels(srcIndex)
            AtlasTexture.m_Pixels(dstIndex + 1) = tex.m_Pixels(srcIndex + 1)
            AtlasTexture.m_Pixels(dstIndex + 2) = tex.m_Pixels(srcIndex + 2)
            AtlasTexture.m_Pixels(dstIndex + 3) = tex.m_Pixels(srcIndex + 3)

        Next dx
    Next dy

End Sub

' ============================================================
' UPLOAD TO GPU
' ============================================================

Public Sub UploadAtlas()

    If Not AtlasTexture Is Nothing Then
        AtlasTexture.UploadToGPU
    End If

End Sub

' ============================================================
' CLEAR ATLAS
' ============================================================

Public Sub Clear()

    Set m_Regions = New Collection

    m_CurrentX = 0
    m_CurrentY = 0
    m_RowHeight = 0

    If Not AtlasTexture Is Nothing Then
        AtlasTexture.Destroy
    End If

    Set AtlasTexture = New Texture2D
    AtlasTexture.CreateEmpty m_AtlasWidth, m_AtlasHeight

End Sub

' ============================================================
' DEBUG
' ============================================================

Public Sub DebugPrint()

    Debug.Print "---- TEXTURE ATLAS v8 ----"
    Debug.Print "Size: "; m_AtlasWidth; "x"; m_AtlasHeight
    Debug.Print "Regions: "; m_Regions.count

End Sub




