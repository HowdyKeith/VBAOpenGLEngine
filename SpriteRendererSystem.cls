Option Explicit

' ============================================================
' SpriteRendererSystem.cls
' ECS 2D Sprite + UI Rendering System (v7 Clean Rewrite)
' Uses Texture Atlas + GPU Batch Renderer
' ============================================================

' =========================
' DEPENDENCIES
' =========================

Public Atlas As TextureAtlasSystem
Public batch As RenderBatcher

' =========================
' SPRITE DATA
' =========================

Private Type Sprite
    x As Single
    y As Single
    w As Single
    h As Single

    UVX As Single
    UVY As Single
    UVW As Single
    UVH As Single

    r As Single
    g As Single
    b As Single
    a As Single

    Layer As Long
End Type

Private m_Sprites() As Sprite
Private m_Count As Long

' =========================
' INIT
' =========================

Private Sub Class_Initialize()

    m_Count = 0
    ReDim m_Sprites(0 To 0)

    Set batch = New RenderBatcher

End Sub

' =========================
' SET ATLAS
' =========================

Public Sub SetAtlas(ByRef atlasSystem As TextureAtlasSystem)
    Set Atlas = atlasSystem
End Sub

' =========================
' CLEAR FRAME BUFFER
' =========================

Public Sub BeginFrame()

    m_Count = 0
    ReDim m_Sprites(0 To 0)

End Sub

' =========================
' ADD SPRITE (CORE API)
' =========================

Public Sub DrawSprite( _
    ByVal textureKey As String, _
    ByVal x As Single, ByVal y As Single, _
    ByVal w As Single, ByVal h As Single, _
    Optional ByVal r As Single = 1, _
    Optional ByVal g As Single = 1, _
    Optional ByVal b As Single = 1, _
    Optional ByVal a As Single = 1, _
    Optional ByVal Layer As Long = 0)

    Dim region As Variant

    If Atlas Is Nothing Then Exit Sub

    region = Atlas.GetRegion(textureKey)

    m_Count = m_Count + 1
    ReDim Preserve m_Sprites(0 To m_Count - 1)

    With m_Sprites(m_Count - 1)

        .x = x
        .y = y
        .w = w
        .h = h

        .UVX = region.x
        .UVY = region.y
        .UVW = region.w
        .UVH = region.h

        .r = r
        .g = g
        .b = b
        .a = a

        .Layer = Layer

    End With

End Sub

' =========================
' SORT BY LAYER (UI SUPPORT)
' =========================

Private Sub SortSprites()

    Dim i As Long, j As Long
    Dim tmp As Sprite

    For i = 0 To m_Count - 2
        For j = i + 1 To m_Count - 1

            If m_Sprites(i).Layer > m_Sprites(j).Layer Then
                tmp = m_Sprites(i)
                m_Sprites(i) = m_Sprites(j)
                m_Sprites(j) = tmp
            End If

        Next j
    Next i

End Sub

' =========================
' RENDER ALL SPRITES
' =========================

Public Sub Render()

    Dim i As Long

    If Atlas Is Nothing Then Exit Sub

    If m_Count = 0 Then Exit Sub

    ' --- ensure correct texture bound ---
    Atlas.Bind 0

    ' --- sort for UI layering ---
    Call SortSprites

    ' --- build batch ---
    For i = 0 To m_Count - 1

        Call batch.AddQuad( _
            m_Sprites(i).x, m_Sprites(i).y, _
            m_Sprites(i).w, m_Sprites(i).h, _
            m_Sprites(i).UVX, m_Sprites(i).UVY, _
            m_Sprites(i).UVW, m_Sprites(i).UVH, _
            m_Sprites(i).r, m_Sprites(i).g, m_Sprites(i).b, m_Sprites(i).a)

    Next i

    ' --- flush GPU ---
    Call batch.Flush

End Sub

' =========================
' UI HELPERS (OPTIONAL LAYER)
' =========================

Public Function DrawButton( _
    ByVal textureKey As String, _
    ByVal x As Single, ByVal y As Single, _
    ByVal w As Single, ByVal h As Single, _
    ByVal hovered As Boolean) As Boolean

    If hovered Then
        Call DrawSprite(textureKey, x, y, w, h, 1, 0.85, 0.85, 1, 10)
    Else
        Call DrawSprite(textureKey, x, y, w, h, 1, 1, 1, 1, 10)
    End If

    DrawButton = True

End Function

' =========================
' CLEANUP
' =========================

Public Sub Destroy()

    Set Atlas = Nothing

    Call batch.Destroy

    m_Count = 0
    ReDim m_Sprites(0 To 0)

End Sub





