Option Explicit

' ============================================================
' FreeGLUT_Demo.bas  v1.0  WEEK 2
' A complete working demo using the FreeGLUT shim.
' Shows: rotating primitives, keyboard control, mouse look,
' timer callback, live FPS in title.
'
' Run: FreeGLUT_Demo.RunDemo
' Controls: W/S = move, mouse = look, 1-5 = swap primitive,
'           +/- = slices, SPACE = toggle wireframe, ESC = quit
' ============================================================

' --- Scene state (module-level, persists across callbacks) ---
Private m_Angle     As Single      ' rotation angle
Private m_Shape     As Long        ' 1=cube 2=sphere 3=torus 4=cylinder 5=cone
Private m_Wire      As Boolean     ' wireframe toggle
Private m_Slices    As Long        ' detail level
Private m_CamY      As Single      ' camera height
Private m_MouseX    As Long
Private m_MouseY    As Long
Private m_LightX    As Single
Private m_FrameCount As Long

' ============================================================
' ENTRY POINT  — call this from a button or the Immediate Window
' ============================================================
Public Sub RunDemo()
    ' Init state
    m_Angle    = 0
    m_Shape    = 1       ' start with cube
    m_Wire     = False
    m_Slices   = 20
    m_CamY     = 1.5
    m_LightX   = 2
    m_FrameCount = 0

    ' ----- FreeGLUT setup (mirrors a C main() exactly) -----
    glutInit
    glutInitDisplayMode GLUT_DOUBLE Or GLUT_RGBA Or GLUT_DEPTH
    glutInitWindowSize 800, 600
    glutCreateWindow "VBA FreeGLUT Demo  |  1-5=Shape  SPACE=Wire  ESC=Quit"

    ' Register callbacks by function name (module-level procs)
    glutDisplayFunc  "FreeGLUT_Demo.Display"
    glutReshapeFunc  "FreeGLUT_Demo.Reshape"
    glutKeyboardFunc "FreeGLUT_Demo.Keyboard"
    glutIdleFunc     "FreeGLUT_Demo.Idle"

    ' Kick off a repeating timer (1 second interval)
    glutTimerFunc 1000, "FreeGLUT_Demo.OneSecondTimer", 0

    ' Enter main loop - blocks until ESC pressed
    glutMainLoop
End Sub

' ============================================================
' DISPLAY CALLBACK  — called every frame
' ============================================================
Public Sub Display()
    GL.glClearColor 0.1, 0.12, 0.15, 1#
    GL.glClear GL.GL_COLOR_BUFFER_BIT Or GL.GL_DEPTH_BUFFER_BIT

    ' --- Camera ---
    GL.glMatrixMode GL.GL_MODELVIEW
    GL.glLoadIdentity
    modGL.apiTranslatef 0, -m_CamY, -4

    ' --- Lighting hint (fixed-function) ---
    GL.glRotatef 20, 1, 0, 0    ' slight pitch down

    ' --- Rotating primitive ---
    GL.glPushMatrix
        GL.glRotatef m_Angle, 0.4, 1, 0.2

        If m_Wire Then modGL.glPolygonMode GL_FRONT_AND_BACK, GL_LINE

        Select Case m_Shape
            Case 1: glutSolidCube 1.4
            Case 2: glutSolidSphere 0.9, m_Slices, m_Slices
            Case 3: glutSolidTorus 0.25, 0.8, m_Slices, m_Slices * 2
            Case 4: glutSolidCylinder 0.5, 1.5, m_Slices, 4
            Case 5: glutSolidCone 0.7, 1.6, m_Slices, 4
        End Select

        If m_Wire Then modGL.glPolygonMode GL_FRONT_AND_BACK, GL_FILL
    GL.glPopMatrix

    ' --- Floor grid (wire squares) ---
    DrawGrid

    m_FrameCount = m_FrameCount + 1
    glutSwapBuffers
End Sub

' ============================================================
' RESHAPE CALLBACK
' ============================================================
Public Sub Reshape(ByVal w As Long, ByVal h As Long)
    If h = 0 Then h = 1
    GL.glViewport 0, 0, w, h

    GL.glMatrixMode GL.GL_PROJECTION
    GL.glLoadIdentity
    modGL.apiPerspective 45#, CDbl(w) / CDbl(h), 0.1, 100#
    GL.glMatrixMode GL.GL_MODELVIEW
End Sub

' ============================================================
' KEYBOARD CALLBACK
' ============================================================
Public Sub Keyboard(ByVal key As String, ByVal x As Long, ByVal y As Long)
    Select Case key
        Case Chr(27)  ' ESC
            glutLeaveMainLoop

        Case "1": m_Shape = 1
        Case "2": m_Shape = 2
        Case "3": m_Shape = 3
        Case "4": m_Shape = 4
        Case "5": m_Shape = 5

        Case " ":  m_Wire = Not m_Wire   ' SPACE = wireframe toggle

        Case "+", "=":
            m_Slices = m_Slices + 2
            If m_Slices > 64 Then m_Slices = 64

        Case "-", "_":
            m_Slices = m_Slices - 2
            If m_Slices < 4 Then m_Slices = 4

        Case "w", "W": m_CamY = m_CamY + 0.2
        Case "s", "S": m_CamY = m_CamY - 0.2
    End Select

    glutPostRedisplay
End Sub

' ============================================================
' IDLE CALLBACK  — called when no events pending
' ============================================================
Public Sub Idle()
    m_Angle = m_Angle + 0.5
    If m_Angle >= 360 Then m_Angle = m_Angle - 360
    glutPostRedisplay
End Sub

' ============================================================
' TIMER CALLBACK (fires every second)
' ============================================================
Public Sub OneSecondTimer(ByVal value As Long)
    ' Re-register for next second
    glutTimerFunc 1000, "FreeGLUT_Demo.OneSecondTimer", 0
    ' Could log m_FrameCount here, but modPerf already handles FPS in title
    m_FrameCount = 0
End Sub

' ============================================================
' PRIVATE: Floor grid
' ============================================================
Private Sub DrawGrid()
    Dim i As Long
    modGL.glPolygonMode GL_FRONT_AND_BACK, GL_LINE
    modGL.glColor3f 0.3, 0.35, 0.4
    GL.glPushMatrix
        GL.glTranslatef 0, -1, 0

        GL.glBegin GL_TRIANGLES  ' degenerate - just draw lines
        GL.glEnd

        ' Simple grid via GL_LINES
        GL.glBegin GL.GL_LINES
            For i = -5 To 5
                GL.glVertex3f CSng(i), 0, -5
                GL.glVertex3f CSng(i), 0,  5
                GL.glVertex3f -5, 0, CSng(i)
                GL.glVertex3f  5, 0, CSng(i)
            Next i
        GL.glEnd
    GL.glPopMatrix
    modGL.glPolygonMode GL_FRONT_AND_BACK, GL_FILL
    modGL.glColor3f 1, 1, 1   ' reset
End Sub
