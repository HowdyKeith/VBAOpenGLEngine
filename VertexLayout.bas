Option Explicit

' ============================================================
' Module: VertexLayout.bas
' Version: v2.0 (2026-04-14)
' Description: Defines all supported vertex layouts + helper constants.
'              Used by MeshBuffer and future shaders.
' ============================================================

Public Enum VertexLayoutType
    Layout_Position             ' 3 floats: XYZ
    Layout_PositionNormal       ' 6 floats: XYZ + Normal
    Layout_PositionNormalUV     ' 8 floats: XYZ + Normal + UV
    Layout_PositionColor        ' 7 floats: XYZ + RGBA (float4)
    ' Add more here as needed (e.g. Layout_PositionNormalUVTangent)
End Enum

' Layout metadata (stride in floats, not bytes)
Public Function GetStride(ByVal layout As VertexLayoutType) As Long
    Select Case layout
        Case Layout_Position:           GetStride = 3
        Case Layout_PositionNormal:     GetStride = 6
        Case Layout_PositionNormalUV:   GetStride = 8
        Case Layout_PositionColor:      GetStride = 7
        Case Else:                      GetStride = 3
    End Select
End Function

' Byte stride (for glVertexAttribPointer)
Public Function GetByteStride(ByVal layout As VertexLayoutType) As Long
    GetByteStride = GetStride(layout) * 4   ' Single = 4 bytes
End Function

' Attribute pointers - call these after binding the VAO + VBO
Public Sub ApplyLayout(ByVal layout As VertexLayoutType)
    Dim stride As Long
    stride = GetByteStride(layout)
    
    Select Case layout
        Case Layout_Position
            GL.glEnableVertexAttribArray 0
            GL.glVertexAttribPointer 0, 3, GL_FLOAT, 0, stride, 0
            
        Case Layout_PositionNormal
            GL.glEnableVertexAttribArray 0   ' Position
            GL.glVertexAttribPointer 0, 3, GL_FLOAT, 0, stride, 0
            GL.glEnableVertexAttribArray 1   ' Normal
            GL.glVertexAttribPointer 1, 3, GL_FLOAT, 0, stride, 12   ' offset after 3 floats
            
        Case Layout_PositionNormalUV
            GL.glEnableVertexAttribArray 0   ' Position
            GL.glVertexAttribPointer 0, 3, GL_FLOAT, 0, stride, 0
            GL.glEnableVertexAttribArray 1   ' Normal
            GL.glVertexAttribPointer 1, 3, GL_FLOAT, 0, stride, 12
            GL.glEnableVertexAttribArray 2   ' UV
            GL.glVertexAttribPointer 2, 2, GL_FLOAT, 0, stride, 24   ' offset after 6 floats
            
        Case Layout_PositionColor
            GL.glEnableVertexAttribArray 0   ' Position
            GL.glVertexAttribPointer 0, 3, GL_FLOAT, 0, stride, 0
            GL.glEnableVertexAttribArray 1   ' Color (RGBA)
            GL.glVertexAttribPointer 1, 4, GL_FLOAT, 0, stride, 12
            
    End Select
End Sub
