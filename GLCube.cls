Option Explicit

' ============================================================
' GLCube.v8.15.cls
' Procedural Cube Primitive (VBA OpenGL Engine Core)
' ============================================================

' --- GEOMETRY STORAGE ---
Private m_Size As Single
Private m_Vertices(0 To 7) As Vector3

' --- STATE ---
Public Active As Boolean
Public name As String

' --- MATERIAL ---
Public MaterialRef As material ' Ensure Material class exists in project

' --- TRANSFORM ---
Public Position As Vector3
Public Rotation As Vector3
Public ScaleVec As Vector3

' =========================
' INIT / TERMINATE
' =========================

Private Sub Class_Initialize()
    m_Size = 1#

    Set Position = New Vector3
    Set Rotation = New Vector3
    Set ScaleVec = New Vector3
    
    ' Default Scale
    ScaleVec.SetValues 1, 1, 1

    Active = True
    name = "Cube"

    ' Initial vertex generation
    GenerateCube
End Sub

Private Sub Class_Terminate()
    ' Cleanup object references to prevent memory leaks
    Set Position = Nothing
    Set Rotation = Nothing
    Set ScaleVec = Nothing
    
    Dim i As Long
    For i = 0 To 7
        Set m_Vertices(i) = Nothing
    Next i
End Sub

' ============================================================
' CUBE GENERATION
' ============================================================

Public Sub SetSize(ByVal size As Single)
    m_Size = size
    Call GenerateCube
End Sub

Private Sub GenerateCube()
    Dim s As Single
    s = m_Size / 2

    ' Initialize vertices if they don't exist
    Dim i As Long
    For i = 0 To 7
        If m_Vertices(i) Is Nothing Then Set m_Vertices(i) = New Vector3
    Next i

    ' -------------------------
    ' 8 VERTICES (Fixed to .SetValues)
    ' -------------------------
    m_Vertices(0).SetValues -s, -s, -s
    m_Vertices(1).SetValues s, -s, -s
    m_Vertices(2).SetValues s, s, -s
    m_Vertices(3).SetValues -s, s, -s

    m_Vertices(4).SetValues -s, -s, s
    m_Vertices(5).SetValues s, -s, s
    m_Vertices(6).SetValues s, s, s
    m_Vertices(7).SetValues -s, s, s
End Sub

' ============================================================
' DRAW CUBE (IMMEDIATE MODE)
' ============================================================

Public Sub Draw()
    If Not Active Then Exit Sub

    ' Apply Material if assigned
    If Not MaterialRef Is Nothing Then
        MaterialRef.Apply Nothing
    End If

    ' --- TRANSFORM BLOCK ---
    GL.glPushMatrix

        GL.glTranslatef Position.x, Position.y, Position.z

        GL.glRotatef Rotation.x, 1, 0, 0
        GL.glRotatef Rotation.y, 0, 1, 0
        GL.glRotatef Rotation.z, 0, 0, 1

        GL.glScalef ScaleVec.x, ScaleVec.y, ScaleVec.z

        ' --- RENDER BLOCK ---
        GL.glBegin GL_QUADS

            ' FRONT
            GL.glNormal3f 0, 0, 1
            GL.glVertex3f m_Vertices(4).x, m_Vertices(4).y, m_Vertices(4).z
            GL.glVertex3f m_Vertices(5).x, m_Vertices(5).y, m_Vertices(5).z
            GL.glVertex3f m_Vertices(6).x, m_Vertices(6).y, m_Vertices(6).z
            GL.glVertex3f m_Vertices(7).x, m_Vertices(7).y, m_Vertices(7).z

            ' BACK
            GL.glNormal3f 0, 0, -1
            GL.glVertex3f m_Vertices(1).x, m_Vertices(1).y, m_Vertices(1).z
            GL.glVertex3f m_Vertices(0).x, m_Vertices(0).y, m_Vertices(0).z
            GL.glVertex3f m_Vertices(3).x, m_Vertices(3).y, m_Vertices(3).z
            GL.glVertex3f m_Vertices(2).x, m_Vertices(2).y, m_Vertices(2).z

            ' LEFT
            GL.glNormal3f -1, 0, 0
            GL.glVertex3f m_Vertices(0).x, m_Vertices(0).y, m_Vertices(0).z
            GL.glVertex3f m_Vertices(4).x, m_Vertices(4).y, m_Vertices(4).z
            GL.glVertex3f m_Vertices(7).x, m_Vertices(7).y, m_Vertices(7).z
            GL.glVertex3f m_Vertices(3).x, m_Vertices(3).y, m_Vertices(3).z

            ' RIGHT
            GL.glNormal3f 1, 0, 0
            GL.glVertex3f m_Vertices(5).x, m_Vertices(5).y, m_Vertices(5).z
            GL.glVertex3f m_Vertices(1).x, m_Vertices(1).y, m_Vertices(1).z
            GL.glVertex3f m_Vertices(2).x, m_Vertices(2).y, m_Vertices(2).z
            GL.glVertex3f m_Vertices(6).x, m_Vertices(6).y, m_Vertices(6).z

            ' TOP
            GL.glNormal3f 0, 1, 0
            GL.glVertex3f m_Vertices(3).x, m_Vertices(3).y, m_Vertices(3).z
            GL.glVertex3f m_Vertices(7).x, m_Vertices(7).y, m_Vertices(7).z
            GL.glVertex3f m_Vertices(6).x, m_Vertices(6).y, m_Vertices(6).z
            GL.glVertex3f m_Vertices(2).x, m_Vertices(2).y, m_Vertices(2).z

            ' BOTTOM
            GL.glNormal3f 0, -1, 0
            GL.glVertex3f m_Vertices(0).x, m_Vertices(0).y, m_Vertices(0).z
            GL.glVertex3f m_Vertices(1).x, m_Vertices(1).y, m_Vertices(1).z
            GL.glVertex3f m_Vertices(5).x, m_Vertices(5).y, m_Vertices(5).z
            GL.glVertex3f m_Vertices(4).x, m_Vertices(4).y, m_Vertices(4).z

        GL.glEnd

    GL.glPopMatrix
End Sub

' ============================================================
' TRANSFORM HELPERS
' ============================================================

Public Sub SetPosition(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Position.SetValues x, y, z
End Sub

Public Sub SetRotation(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    Rotation.SetValues x, y, z
End Sub

Public Sub SetScale(ByVal x As Single, ByVal y As Single, ByVal z As Single)
    ScaleVec.SetValues x, y, z
End Sub

' ============================================================
' DEBUG
' ============================================================

Public Sub DebugPrint()
    Debug.Print "---- [GLCube: " & name & "] ----"
    Debug.Print "Size: " & m_Size
    Debug.Print "Pos:  " & Position.x & ", " & Position.y & ", " & Position.z
    Debug.Print "Active: " & Active
End Sub


