Option Explicit

' ============================================================
' Class      : frustum
' Version    : v8.11
' Description: Camera Frustum for Culling Operations
' ============================================================

Private m_Planes(0 To 5, 0 To 3) As Single

' Updates the 6 planes based on the current View-Projection matrix
' Pass a 16-element 1D array (0 to 15)
Public Sub ExtractPlanes(ByRef clip() As Single)
    ' Right
    m_Planes(0, 0) = clip(3) - clip(0): m_Planes(0, 1) = clip(7) - clip(4)
    m_Planes(0, 2) = clip(11) - clip(8): m_Planes(0, 3) = clip(15) - clip(12)
    ' Left
    m_Planes(1, 0) = clip(3) + clip(0): m_Planes(1, 1) = clip(7) + clip(4)
    m_Planes(1, 2) = clip(11) + clip(8): m_Planes(1, 3) = clip(15) + clip(12)
    ' Bottom
    m_Planes(2, 0) = clip(3) + clip(1): m_Planes(2, 1) = clip(7) + clip(5)
    m_Planes(2, 2) = clip(11) + clip(9): m_Planes(2, 3) = clip(15) + clip(13)
    ' Top
    m_Planes(3, 0) = clip(3) - clip(1): m_Planes(3, 1) = clip(7) - clip(5)
    m_Planes(3, 2) = clip(11) - clip(9): m_Planes(3, 3) = clip(15) - clip(13)
    ' Far
    m_Planes(4, 0) = clip(3) - clip(2): m_Planes(4, 1) = clip(7) - clip(6)
    m_Planes(4, 2) = clip(11) - clip(10): m_Planes(4, 3) = clip(15) - clip(14)
    ' Near
    m_Planes(5, 0) = clip(3) + clip(2): m_Planes(5, 1) = clip(7) + clip(6)
    m_Planes(5, 2) = clip(11) + clip(10): m_Planes(5, 3) = clip(15) + clip(14)
    
    ' (Optional: Add plane normalization logic here if using distance-based culling)
End Sub

' --- THE VISIBILITY CHECK ---
Public Function IsBoxVisible(ByVal minX As Single, ByVal minY As Single, ByVal minZ As Single, _
                             ByVal maxX As Single, ByVal maxY As Single, ByVal maxZ As Single) As Boolean
    
    Dim i As Long
    For i = 0 To 5
        If m_Planes(i, 0) * (IIf(m_Planes(i, 0) > 0, maxX, minX)) + _
           m_Planes(i, 1) * (IIf(m_Planes(i, 1) > 0, maxY, minY)) + _
           m_Planes(i, 2) * (IIf(m_Planes(i, 2) > 0, maxZ, minZ)) + _
           m_Planes(i, 3) < 0 Then
           IsBoxVisible = False
           Exit Function
        End If
    Next i
    IsBoxVisible = True
End Function




