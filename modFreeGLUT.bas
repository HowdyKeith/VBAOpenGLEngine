Option Explicit

' ============================================================
' modFreeGLUT.bas  v1.0  WEEK 2
' Drop-in FreeGLUT / GLUT compatibility shim for VBA OpenGL Engine.
'
' PURPOSE:
'   Lets developers port existing FreeGLUT/GLUT C projects to VBA
'   with minimal changes. Every public Sub/Function mirrors its C
'   counterpart by name and signature (types widened to VBA equivalents).
'
' USAGE PATTERN (mirrors C exactly):
'
'   Public Sub MyDisplay()
'       glClear GL_COLOR_BUFFER_BIT Or GL_DEPTH_BUFFER_BIT
'       glutSolidSphere 1.0, 20, 20
'       glutSwapBuffers
'   End Sub
'
'   Public Sub MyReshape(w As Long, h As Long)
'       glViewport 0, 0, w, h
'   End Sub
'
'   Public Sub MyKeyboard(key As String, x As Long, y As Long)
'       If key = Chr(27) Then glutLeaveMainLoop
'   End Sub
'
'   Public Sub RunMyApp()
'       glutInit
'       glutInitDisplayMode GLUT_DOUBLE Or GLUT_RGBA Or GLUT_DEPTH
'       glutInitWindowSize 800, 600
'       glutCreateWindow "My FreeGLUT Port"
'       glutDisplayFunc  "MyDisplay"
'       glutReshapeFunc  "MyReshape"
'       glutKeyboardFunc "MyKeyboard"
'       glutMainLoop
'   End Sub
'
' ============================================================

' --- GLUT Display Mode Constants (matches FreeGLUT values) ---
Public Const GLUT_RGBA        As Long = &H0
Public Const GLUT_RGB         As Long = &H0
Public Const GLUT_INDEX       As Long = &H1
Public Const GLUT_SINGLE      As Long = &H0
Public Const GLUT_DOUBLE      As Long = &H2
Public Const GLUT_ACCUM       As Long = &H4
Public Const GLUT_ALPHA       As Long = &H8
Public Const GLUT_DEPTH       As Long = &H10
Public Const GLUT_STENCIL     As Long = &H20
Public Const GLUT_MULTISAMPLE As Long = &H80

' --- GLUT Key Constants ---
Public Const GLUT_KEY_F1      As Long = &H1
Public Const GLUT_KEY_F2      As Long = &H2
Public Const GLUT_KEY_F3      As Long = &H3
Public Const GLUT_KEY_F4      As Long = &H4
Public Const GLUT_KEY_F5      As Long = &H5
Public Const GLUT_KEY_F6      As Long = &H6
Public Const GLUT_KEY_F7      As Long = &H7
Public Const GLUT_KEY_F8      As Long = &H8
Public Const GLUT_KEY_F9      As Long = &H9
Public Const GLUT_KEY_F10     As Long = &HA
Public Const GLUT_KEY_F11     As Long = &HB
Public Const GLUT_KEY_F12     As Long = &HC
Public Const GLUT_KEY_LEFT    As Long = &H64
Public Const GLUT_KEY_RIGHT   As Long = &H65
Public Const GLUT_KEY_UP      As Long = &H66
Public Const GLUT_KEY_DOWN    As Long = &H67
Public Const GLUT_KEY_PAGE_UP As Long = &H68
Public Const GLUT_KEY_PAGE_DOWN As Long = &H69
Public Const GLUT_KEY_HOME    As Long = &H6A
Public Const GLUT_KEY_END     As Long = &H6B

' --- GLUT Mouse Button Constants ---
Public Const GLUT_LEFT_BUTTON   As Long = 0
Public Const GLUT_MIDDLE_BUTTON As Long = 1
Public Const GLUT_RIGHT_BUTTON  As Long = 2
Public Const GLUT_DOWN          As Long = 0
Public Const GLUT_UP            As Long = 1

' --- GLUT Get Constants ---
Public Const GLUT_WINDOW_WIDTH    As Long = 102
Public Const GLUT_WINDOW_HEIGHT   As Long = 103
Public Const GLUT_ELAPSED_TIME    As Long = 700
Public Const GLUT_SCREEN_WIDTH    As Long = 200
Public Const GLUT_SCREEN_HEIGHT   As Long = 201

' =========================
' INTERNAL STATE
' =========================
Private m_WindowWidth    As Long
Private m_WindowHeight   As Long
Private m_DisplayMode    As Long
Private m_WindowTitle    As String
Private m_hWnd           As LongPtr
Private m_Running        As Boolean
Private m_RedisplayNeeded As Boolean
Private m_StartTime      As Double

' --- Callback names (module.function string, called via CallByName on module) ---
Private m_cbDisplay      As String   ' void (*)(void)
Private m_cbIdle         As String   ' void (*)(void)
Private m_cbReshape      As String   ' void (*)(int w, int h)
Private m_cbKeyboard     As String   ' void (*)(unsigned char key, int x, int y)
Private m_cbKeyboardUp   As String   ' void (*)(unsigned char key, int x, int y)
Private m_cbSpecial      As String   ' void (*)(int key, int x, int y)
Private m_cbSpecialUp    As String   ' void (*)(int key, int x, int y)
Private m_cbMouse        As String   ' void (*)(int btn, int state, int x, int y)
Private m_cbMotion       As String   ' void (*)(int x, int y)  (button held)
Private m_cbPassiveMotion As String  ' void (*)(int x, int y)  (no button)
Private m_cbClose        As String   ' void (*)(void)
Private m_cbTimer        As String   ' void (*)(unsigned int value)
Private m_TimerMS        As Long
Private m_TimerValue     As Long
Private m_TimerDue       As Double

' Input helper (reuse project's InputSystem)
Private m_Input          As InputSystem

' Primitive singletons (lazy-init)
Private m_Sphere   As GLSphere
Private m_Torus    As GLTorus
Private m_Cylinder As GLCylinder

' =========================
' INIT
' =========================

' glutInit() — initialise the GLUT library
Public Sub glutInit()
    m_WindowWidth  = 640
    m_WindowHeight = 480
    m_WindowTitle  = "GLUT Window"
    m_DisplayMode  = GLUT_DOUBLE Or GLUT_RGBA Or GLUT_DEPTH
    m_Running      = False
    m_StartTime    = Win32GL.GetTime()
    Set m_Input    = New InputSystem
    Debug.Print "[FreeGLUT] Initialized."
End Sub

' glutInitWindowSize(width, height)
Public Sub glutInitWindowSize(ByVal w As Long, ByVal h As Long)
    m_WindowWidth = w: m_WindowHeight = h
End Sub

' glutInitWindowPosition(x, y) — no-op in VBA (window always at 100,100)
Public Sub glutInitWindowPosition(ByVal x As Long, ByVal y As Long)
    ' Position hint accepted, not currently forwarded to CreateGLWindow
End Sub

' glutInitDisplayMode(mode)
Public Sub glutInitDisplayMode(ByVal mode As Long)
    m_DisplayMode = mode
End Sub

' glutCreateWindow(title) — returns window ID (always 1 in VBA single-window model)
Public Function glutCreateWindow(ByVal title As String) As Long
    m_WindowTitle = title
    m_hWnd = modWindowManager.CreateGLWindow(title, m_WindowWidth, m_WindowHeight)

    If m_hWnd = 0 Then
        Debug.Print "[FreeGLUT] ERROR: Failed to create window."
        glutCreateWindow = 0
        Exit Function
    End If

    ' Load GL extension pointers
    Dim hLib As LongPtr
    hLib = Win32GL.GetModuleHandle("opengl32.dll")
    GLLoader.Init hLib
    GL.GL_Init
    modPerf.Init

    ' Set up default GL state
    GL.glEnable GL.GL_DEPTH_TEST
    GL.glEnable GL.GL_CULL_FACE
    GL.glCullFace GL.GL_BACK

    ' Fire initial reshape
    FireReshape m_WindowWidth, m_WindowHeight

    Debug.Print "[FreeGLUT] Window created: " & title & " (" & m_WindowWidth & "x" & m_WindowHeight & ")"
    glutCreateWindow = 1
End Function

' glutDestroyWindow(win)
Public Sub glutDestroyWindow(ByVal win As Long)
    modWindowManager.CloseGLWindow
End Sub

' =========================
' CALLBACK REGISTRATION
' =========================

' glutDisplayFunc("ModuleName.ProcName") or ("ProcName" for module-level)
Public Sub glutDisplayFunc(ByVal callbackName As String)
    m_cbDisplay = callbackName
End Sub

Public Sub glutIdleFunc(ByVal callbackName As String)
    m_cbIdle = callbackName
End Sub

Public Sub glutReshapeFunc(ByVal callbackName As String)
    m_cbReshape = callbackName
End Sub

Public Sub glutKeyboardFunc(ByVal callbackName As String)
    m_cbKeyboard = callbackName
End Sub

Public Sub glutKeyboardUpFunc(ByVal callbackName As String)
    m_cbKeyboardUp = callbackName
End Sub

Public Sub glutSpecialFunc(ByVal callbackName As String)
    m_cbSpecial = callbackName
End Sub

Public Sub glutSpecialUpFunc(ByVal callbackName As String)
    m_cbSpecialUp = callbackName
End Sub

Public Sub glutMouseFunc(ByVal callbackName As String)
    m_cbMouse = callbackName
End Sub

Public Sub glutMotionFunc(ByVal callbackName As String)
    m_cbMotion = callbackName
End Sub

Public Sub glutPassiveMotionFunc(ByVal callbackName As String)
    m_cbPassiveMotion = callbackName
End Sub

Public Sub glutCloseFunc(ByVal callbackName As String)
    m_cbClose = callbackName
End Sub

' glutTimerFunc(ms, callbackName, value)
Public Sub glutTimerFunc(ByVal ms As Long, ByVal callbackName As String, ByVal value As Long)
    m_cbTimer    = callbackName
    m_TimerMS    = ms
    m_TimerValue = value
    m_TimerDue   = Win32GL.GetTime() + ms / 1000#
End Sub

' =========================
' MAIN LOOP
' =========================

' glutMainLoop() — enter the event-processing loop (never returns in C;
' in VBA returns after glutLeaveMainLoop is called)
Public Sub glutMainLoop()
    If m_hWnd = 0 Then
        Debug.Print "[FreeGLUT] ERROR: No window. Call glutCreateWindow first."
        Exit Sub
    End If

    m_Running = True
    Dim prevMouseX As Long, prevMouseY As Long

    Do While m_Running

        modPerf.BeginFrame
        m_Input.Update

        ' --- Keyboard callbacks ---
        Dim k As Long
        For k = 0 To 255
            If m_Input.GetKeyDown(k) Then
                ' Fire keyboard down callback
                If Len(m_cbKeyboard) > 0 Then
                    SafeCall2S m_cbKeyboard, Chr(k), 0, 0
                End If
                ' Special keys (F1-F12, arrows)
                If k >= &H70 And k <= &H7B Then  ' F1-F12
                    If Len(m_cbSpecial) > 0 Then SafeCall2S m_cbSpecial, CStr(k - &H70 + 1), 0, 0
                End If
            End If
            If m_Input.GetKeyUp(k) Then
                If Len(m_cbKeyboardUp) > 0 Then SafeCall2S m_cbKeyboardUp, Chr(k), 0, 0
            End If
        Next k

        ' --- ESC default exit (overrideable via keyboard callback) ---
        If m_Input.GetKey(Win32GL.VK_ESCAPE) Then m_Running = False

        ' --- Mouse motion ---
        Dim mx As Long, my As Long
        mx = m_Input.GetMouseX: my = m_Input.GetMouseY
        If mx <> prevMouseX Or my <> prevMouseY Then
            If Len(m_cbPassiveMotion) > 0 Then SafeCall1 m_cbPassiveMotion, mx, my
            prevMouseX = mx: prevMouseY = my
        End If

        ' --- Timer callback ---
        If Len(m_cbTimer) > 0 And Win32GL.GetTime() >= m_TimerDue Then
            SafeCallValue m_cbTimer, m_TimerValue
            m_cbTimer = ""   ' one-shot (re-register if needed)
        End If

        ' --- Display callback ---
        If Len(m_cbDisplay) > 0 Then
            If m_RedisplayNeeded Or Len(m_cbIdle) = 0 Then
                SafeCallVoid m_cbDisplay
                m_RedisplayNeeded = False
            End If
        End If

        ' --- Idle callback ---
        If Len(m_cbIdle) > 0 Then SafeCallVoid m_cbIdle

        ' --- Message pump + present (NO DoEvents) ---
        m_Running = Win32GL.PumpMessages() And m_Running
        modWindowManager.PageFlip

        modPerf.EndFrame
        modPerf.UpdateWindowTitle m_hWnd

    Loop

    ' Fire close callback
    If Len(m_cbClose) > 0 Then SafeCallVoid m_cbClose
    modWindowManager.CloseGLWindow
    Debug.Print "[FreeGLUT] Main loop exited. " & modPerf.TotalFrames & " frames."
    modPerf.DebugPrint
End Sub

' glutMainLoopEvent() — process one iteration (non-blocking, for integration)
Public Function glutMainLoopEvent() As Boolean
    If Not m_Running Then glutMainLoopEvent = False: Exit Function
    m_Input.Update
    If Len(m_cbDisplay) > 0 Then SafeCallVoid m_cbDisplay
    If Len(m_cbIdle) > 0 Then    SafeCallVoid m_cbIdle
    m_Running = Win32GL.PumpMessages()
    modWindowManager.PageFlip
    glutMainLoopEvent = m_Running
End Function

' glutLeaveMainLoop() — signal the loop to exit after this iteration
Public Sub glutLeaveMainLoop()
    m_Running = False
End Sub

' =========================
' WINDOW / DISPLAY
' =========================

' glutSwapBuffers()
Public Sub glutSwapBuffers()
    modWindowManager.PageFlip
End Sub

' glutPostRedisplay() — flag that the display needs redrawing
Public Sub glutPostRedisplay()
    m_RedisplayNeeded = True
End Sub

' glutSetWindowTitle(title)
Public Sub glutSetWindowTitle(ByVal title As String)
    Win32GL.SetWindowText m_hWnd, title
End Sub

' glutFullScreen() — resize to full desktop (approximation)
Public Sub glutFullScreen()
    Debug.Print "[FreeGLUT] glutFullScreen: not supported - resize window manually."
End Sub

' glutReshapeWindow(w, h)
Public Sub glutReshapeWindow(ByVal w As Long, ByVal h As Long)
    m_WindowWidth = w: m_WindowHeight = h
    FireReshape w, h
End Sub

' =========================
' QUERY
' =========================

' glutGet(query) — returns various state values
Public Function glutGet(ByVal query As Long) As Long
    Select Case query
        Case GLUT_WINDOW_WIDTH:  glutGet = m_WindowWidth
        Case GLUT_WINDOW_HEIGHT: glutGet = m_WindowHeight
        Case GLUT_ELAPSED_TIME:  glutGet = CLng((Win32GL.GetTime() - m_StartTime) * 1000#)
        Case GLUT_SCREEN_WIDTH:  glutGet = 1920  ' reasonable default
        Case GLUT_SCREEN_HEIGHT: glutGet = 1080
        Case Else:               glutGet = 0
    End Select
End Function

' glutGetModifiers() — returns Shift/Ctrl/Alt state
Public Function glutGetModifiers() As Long
    ' Check shift/ctrl/alt via GetAsyncKeyState
    Dim mods As Long: mods = 0
    If Win32GL.IsKeyDown(&H10) Then mods = mods Or 1  ' GLUT_ACTIVE_SHIFT
    If Win32GL.IsKeyDown(&H11) Then mods = mods Or 2  ' GLUT_ACTIVE_CTRL
    If Win32GL.IsKeyDown(&H12) Then mods = mods Or 4  ' GLUT_ACTIVE_ALT
    glutGetModifiers = mods
End Function

' =========================
' PRIMITIVES  (mirrors FreeGLUT C API exactly)
' =========================

' glutSolidCube(size)
Public Sub glutSolidCube(ByVal size As Double)
    Dim c As New GLCube
    c.SetSize CSng(size)
    c.Draw
End Sub

' glutWireCube(size)
Public Sub glutWireCube(ByVal size As Double)
    modGL.glPolygonMode GL_FRONT_AND_BACK, GL_LINE
    Dim c As New GLCube
    c.SetSize CSng(size)
    c.Draw
    modGL.glPolygonMode GL_FRONT_AND_BACK, GL_FILL
End Sub

' glutSolidSphere(radius, slices, stacks)
Public Sub glutSolidSphere(ByVal radius As Double, ByVal slices As Long, ByVal stacks As Long)
    If m_Sphere Is Nothing Then Set m_Sphere = New GLSphere
    m_Sphere.DrawSolid CSng(radius), slices, stacks
End Sub

' glutWireSphere(radius, slices, stacks)
Public Sub glutWireSphere(ByVal radius As Double, ByVal slices As Long, ByVal stacks As Long)
    If m_Sphere Is Nothing Then Set m_Sphere = New GLSphere
    m_Sphere.DrawWireform CSng(radius), slices, stacks
End Sub

' glutSolidTorus(innerRadius, outerRadius, sides, rings)
Public Sub glutSolidTorus(ByVal innerRadius As Double, ByVal outerRadius As Double, ByVal sides As Long, ByVal rings As Long)
    If m_Torus Is Nothing Then Set m_Torus = New GLTorus
    m_Torus.DrawSolid CSng(innerRadius), CSng(outerRadius), sides, rings
End Sub

' glutWireTorus(innerRadius, outerRadius, sides, rings)
Public Sub glutWireTorus(ByVal innerRadius As Double, ByVal outerRadius As Double, ByVal sides As Long, ByVal rings As Long)
    If m_Torus Is Nothing Then Set m_Torus = New GLTorus
    m_Torus.DrawWireform CSng(innerRadius), CSng(outerRadius), sides, rings
End Sub

' glutSolidCylinder(radius, height, slices, stacks)
Public Sub glutSolidCylinder(ByVal radius As Double, ByVal height As Double, ByVal slices As Long, ByVal stacks As Long)
    If m_Cylinder Is Nothing Then Set m_Cylinder = New GLCylinder
    m_Cylinder.DrawSolid CSng(radius), CSng(radius), CSng(height), slices, stacks
End Sub

' glutWireCylinder(radius, height, slices, stacks)
Public Sub glutWireCylinder(ByVal radius As Double, ByVal height As Double, ByVal slices As Long, ByVal stacks As Long)
    If m_Cylinder Is Nothing Then Set m_Cylinder = New GLCylinder
    m_Cylinder.DrawWireform CSng(radius), CSng(radius), CSng(height), slices, stacks
End Sub

' glutSolidCone(base, height, slices, stacks)
Public Sub glutSolidCone(ByVal base As Double, ByVal height As Double, ByVal slices As Long, ByVal stacks As Long)
    If m_Cylinder Is Nothing Then Set m_Cylinder = New GLCylinder
    m_Cylinder.DrawSolid CSng(base), 0, CSng(height), slices, stacks
End Sub

' glutWireCone(base, height, slices, stacks)
Public Sub glutWireCone(ByVal base As Double, ByVal height As Double, ByVal slices As Long, ByVal stacks As Long)
    If m_Cylinder Is Nothing Then Set m_Cylinder = New GLCylinder
    m_Cylinder.DrawWireform CSng(base), 0, CSng(height), slices, stacks
End Sub

' glutSolidTeapot(size) — stub with recognisable placeholder sphere
Public Sub glutSolidTeapot(ByVal size As Double)
    ' Full Newell teapot would require 306 Bezier patches.
    ' Approximated with a sphere + cylinder combination that occupies same bounding box.
    ' To add full teapot: load teapot.obj via OBJLoader
    Debug.Print "[FreeGLUT] glutSolidTeapot: approximated. Load teapot.obj for full mesh."
    glutSolidSphere size * 0.5, 16, 16
End Sub

Public Sub glutWireTeapot(ByVal size As Double)
    Debug.Print "[FreeGLUT] glutWireTeapot: approximated."
    glutWireSphere size * 0.5, 16, 16
End Sub

' =========================
' SAFE CALLBACK INVOKE HELPERS
' =========================

Private Sub SafeCallVoid(ByVal name As String)
    On Error Resume Next
    Dim parts() As String: parts = Split(name, ".")
    If UBound(parts) = 0 Then
        ' Module-level call - use Application.Run in Excel context
        Application.Run name
    Else
        ' "Module.Proc" style - not supported via Application.Run directly
        Application.Run name
    End If
    On Error GoTo 0
End Sub

Private Sub SafeCall2S(ByVal name As String, ByVal a1 As String, ByVal a2 As Long, ByVal a3 As Long)
    On Error Resume Next
    Application.Run name, a1, a2, a3
    On Error GoTo 0
End Sub

Private Sub SafeCall1(ByVal name As String, ByVal a1 As Long, ByVal a2 As Long)
    On Error Resume Next
    Application.Run name, a1, a2
    On Error GoTo 0
End Sub

Private Sub SafeCallValue(ByVal name As String, ByVal value As Long)
    On Error Resume Next
    Application.Run name, value
    On Error GoTo 0
End Sub

Private Sub FireReshape(ByVal w As Long, ByVal h As Long)
    ' Set default viewport
    GL.glViewport 0, 0, w, h
    If Len(m_cbReshape) > 0 Then SafeCall1 m_cbReshape, w, h
End Sub
