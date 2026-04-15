Option Explicit

Private m_Direction As Vector3
Private m_ShadowMap() As Single
Private m_Res As Long

Private Sub Class_Initialize()
    Set m_Direction = New Vector3
    m_Res = 256
    ReDim m_ShadowMap(0 To m_Res * m_Res - 1)
End Sub

Public Sub SetDirection(x As Single, y As Single, z As Single)
    m_Direction.x = x
    m_Direction.y = y
    m_Direction.z = z
End Sub

Public Sub ClearShadowMap()

    Dim i As Long
    For i = 0 To UBound(m_ShadowMap)
        m_ShadowMap(i) = 1E+20
    Next i

End Sub

Public Sub WriteDepth(ByVal x As Long, ByVal y As Long, ByVal d As Single)

    Dim i As Long
    i = x + y * m_Res

    If i >= 0 And i <= UBound(m_ShadowMap) Then
        If d < m_ShadowMap(i) Then m_ShadowMap(i) = d
    End If

End Sub

Public Function SampleDepth(ByVal x As Long, ByVal y As Long) As Single

    Dim i As Long
    i = x + y * m_Res

    If i < 0 Or i > UBound(m_ShadowMap) Then
        SampleDepth = 1E+20
    Else
        SampleDepth = m_ShadowMap(i)
    End If

End Function




