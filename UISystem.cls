Option Explicit

' ============================================================
' UISystem.cls
' Immediate Mode UI (Sprite-based)
' VBA OpenGL v7 ECS Engine
' ============================================================

' =========================
' DEPENDENCIES
' =========================

Public SpriteRenderer As SpriteRendererSystem

' =========================
' UI ELEMENT STATE
' =========================

Private Type UIButton
    x As Single
    y As Single
    w As Single
    h As Single

    text As String
    textureKey As String

    Hover As Boolean
    clicked As Boolean
End Type

Private m_Buttons() As UIButton
Private m_ButtonCount As Long

' =========================
' INIT
' =========================

Private Sub Class_Initialize()
    m_ButtonCount = 0
    ReDim m_Buttons(0 To 0)
End Sub

' =========================
' SET SPRITE RENDERER
' =========================

Public Sub SetRenderer(ByRef renderer As SpriteRendererSystem)
    Set SpriteRenderer = renderer
End Sub

' =========================
' CREATE BUTTON
' =========================

Public Function CreateButton( _
    ByVal x As Single, ByVal y As Single, _
    ByVal w As Single, ByVal h As Single, _
    ByVal text As String, _
    ByVal textureKey As String) As Long

    m_ButtonCount = m_ButtonCount + 1
    ReDim Preserve m_Buttons(0 To m_ButtonCount - 1)

    With m_Buttons(m_ButtonCount - 1)
        .x = x
        .y = y
        .w = w
        .h = h
        .text = text
        .textureKey = textureKey
        .Hover = False
        .clicked = False
    End With

    CreateButton = m_ButtonCount - 1

End Function

' =========================
' UPDATE UI (INPUT LOGIC HOOK)
' =========================

Public Sub UpdateMouse(ByVal mx As Single, ByVal my As Single, ByVal clicked As Boolean)

    Dim i As Long

    For i = 0 To m_ButtonCount - 1

        With m_Buttons(i)

            .Hover = (mx >= .x And mx <= .x + .w And _
                      my >= .y And my <= .y + .h)

            If .Hover And clicked Then
                .clicked = True
            Else
                .clicked = False
            End If

        End With

    Next i

End Sub

' =========================
' RENDER UI
' =========================

Public Sub Render()

    Dim i As Long

    If SpriteRenderer Is Nothing Then Exit Sub

    For i = 0 To m_ButtonCount - 1

        With m_Buttons(i)

            If .Hover Then

                Call SpriteRenderer.DrawSprite( _
                    .textureKey, .x, .y, .w, .h, _
                    1, 0.8, 0.8, 1)

            Else

                Call SpriteRenderer.DrawSprite( _
                    .textureKey, .x, .y, .w, .h, _
                    1, 1, 1, 1)

            End If

        End With

    Next i

End Sub

' =========================
' GET BUTTON STATE
' =========================

Public Function IsClicked(ByVal index As Long) As Boolean

    If index < 0 Or index >= m_ButtonCount Then Exit Function

    IsClicked = m_Buttons(index).clicked

End Function





