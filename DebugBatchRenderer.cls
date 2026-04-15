Option Explicit

' =========================================================
' DebugBatchRenderer.v8.cls
' CPU-SIDE ECS BATCH RENDERER (NO OpenGL)
' Used for debugging BVH + SpatialGrid + Physics
' =========================================================

' =========================
' DRAW COMMAND
' =========================

Private Type DrawCommand
    x As Single
    y As Single
    z As Single

    r As Single
    g As Single
    b As Single

    size As Single
    entityID As Long
End Type

' =========================
' STORAGE
' =========================

Private m_Commands() As DrawCommand
Private m_Count As Long

' =========================
' INIT
' =========================

Private Sub Class_Initialize()
    m_Count = 0
    ReDim m_Commands(0 To 0)
End Sub

' =========================================================
' FRAME BEGIN
' =========================================================

Public Sub BeginFrame()
    m_Count = 0
End Sub

' =========================================================
' SUBMIT ENTITY (FROM ECS / BVH / GRID)
' =========================================================

Public Sub Submit(ByVal entityID As Long, _
                  ByVal x As Single, _
                  ByVal y As Single, _
                  ByVal z As Single, _
                  ByVal r As Single, _
                  ByVal g As Single, _
                  ByVal b As Single, _
                  Optional ByVal size As Single = 5!)

    m_Count = m_Count + 1
    ReDim Preserve m_Commands(0 To m_Count - 1)

    With m_Commands(m_Count - 1)
        .entityID = entityID
        .x = x
        .y = y
        .z = z

        .r = r
        .g = g
        .b = b

        .size = size
    End With

End Sub

' =========================================================
' DRAW PASS (DEBUG VISUALIZATION)
' =========================================================

Public Sub Flush()

    Dim i As Long

    For i = 0 To m_Count - 1

        Dim c As DrawCommand
        c = m_Commands(i)

        ' =====================================================
        ' DEBUG OUTPUT (replace later with real GPU layer)
        ' =====================================================

        Call DrawPoint(c.x, c.y, c.z, c.r, c.g, c.b, c.size)

    Next i

End Sub

' =========================================================
' DEBUG DRAW FUNCTION
' =========================================================

Private Sub DrawPoint(ByVal x As Single, ByVal y As Single, ByVal z As Single, _
                      ByVal r As Single, ByVal g As Single, ByVal b As Single, _
                      ByVal size As Single)

    ' Replace this with your renderer hook:
    ' - GDI debug draw
    ' - DirectX overlay
    ' - VBA OpenGL debug viewport

    Debug.Print "DRAW:", x, y, z, "Color:", r, g, b, "Size:", size

End Sub




