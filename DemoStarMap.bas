Option Explicit

' ============================================================
' DemoStarMap.bas  v1.0  WEEK 3
' 3D Interactive Star Map rendered from Excel data.
'
' SETUP:
'   1. Run modExcelBridge.GenerateSampleStarData to create the
'      "Stars" sheet, OR paste your own HYG catalog data with:
'      Col A=Name, B=X, C=Y, D=Z, E=Magnitude, F=ColorIndex(B-V)
'   2. Call DemoStarMap.Run from a button or Immediate Window
'
' CONTROLS:
'   W/S          - fly forward/back
'   A/D          - strafe left/right
'   Q/E          - fly up/down
'   Mouse        - look around
'   +/-          - increase/decrease star brightness
'   C            - cycle colour mode (BV / Spectral class / White)
'   H            - highlight nearest star in Excel sheet
'   ESC          - quit
' ============================================================

Private Const SHEET_NAME    As String = "Stars"
Private Const DATA_START    As Long   = 2      ' row 1 = header
Private Const MAX_STARS     As Long   = 10000  ' cap for VBA performance
Private Const FLOATS_PER_STAR As Long = 7      ' x,y,z,r,g,b,size

' --- GL handles ---
Private m_VAO       As Long
Private m_VBO       As Long
Private m_Shader    As ShaderProgram
Private m_StarCount As Long

' --- Camera ---
Private m_CamX As Single, m_CamY As Single, m_CamZ As Single
Private m_Yaw  As Single, m_Pitch As Single

' --- Runtime state ---
Private m_BrightScale As Single
Private m_ColorMode   As Long     ' 0=BV 1=Spectral 2=White
Private m_Input       As InputSystem
Private m_Running     As Boolean
Private m_HighlightRow As Long     ' currently highlighted star

' Raw data arrays for picking
Private m_StarPosX() As Single
Private m_StarPosY() As Single
Private m_StarPosZ() As Single
Private m_StarNames() As String

' ============================================================
' ENTRY POINT
' ============================================================
Public Sub Run()
    ' Ensure sample data exists
    If Not modExcelBridge.SheetExists(SHEET_NAME) Then
        Debug.Print "[StarMap] Creating sample star data..."
        modExcelBridge.GenerateSampleStarData SHEET_NAME, 3000
    End If

    ' --- Window + GL ---
    Dim hWnd As LongPtr
    hWnd = modWindowManager.CreateGLWindow("VBA Star Map  |  " & SHEET_NAME & " sheet", 1024, 768)
    If hWnd = 0 Then MsgBox "Failed to create GL window", vbCritical: Exit Sub

    Dim hLib As LongPtr
    hLib = Win32GL.GetModuleHandle("opengl32.dll")
    GLLoader.Init hLib
    GL.GL_Init
    modPerf.Init

    ' --- GL state ---
    GL.glEnable GL.GL_DEPTH_TEST
    GL.glEnable GL.GL_BLEND
    GL.glBlendFunc GL.GL_SRC_ALPHA, GL.GL_ONE  ' additive blending for glow
    GL.glEnable GL.GL_PROGRAM_POINT_SIZE       ' let shader control point size

    ' --- Build GPU star buffer ---
    InitStars
    If m_StarCount = 0 Then
        MsgBox "No star data found on sheet '" & SHEET_NAME & "'", vbExclamation
        modWindowManager.CloseGLWindow
        Exit Sub
    End If

    ' --- Camera start ---
    m_CamX = 0: m_CamY = 0: m_CamZ = 50
    m_Yaw = 0: m_Pitch = 0
    m_BrightScale = 1!
    m_ColorMode   = 0
    m_HighlightRow = -1

    Set m_Input = New InputSystem
    m_Input.SetWindowHandle hWnd

    Dim prevMouseX As Long, prevMouseY As Long
    prevMouseX = m_Input.GetMouseX: prevMouseY = m_Input.GetMouseY

    ' --- Main loop ---
    Dim t0 As Double: t0 = Win32GL.GetTime()
    m_Running = True

    Do While m_Running
        modPerf.BeginFrame

        Dim t1 As Double: t1 = Win32GL.GetTime()
        Dim dt As Single: dt = CSng(t1 - t0): t0 = t1
        If dt > 0.05 Then dt = 0.05

        m_Input.Update

        ' --- Input ---
        If m_Input.GetKey(Win32GL.VK_ESCAPE) Then m_Running = False

        ' Movement
        Dim spd As Single: spd = 30 * dt
        Dim fwdX As Single, fwdZ As Single
        fwdX = CSng(Sin(m_Yaw * 3.14159 / 180))
        fwdZ = CSng(-Cos(m_Yaw * 3.14159 / 180))

        If m_Input.GetKey(Win32GL.VK_W) Then m_CamX = m_CamX + fwdX * spd: m_CamZ = m_CamZ + fwdZ * spd
        If m_Input.GetKey(Win32GL.VK_S) Then m_CamX = m_CamX - fwdX * spd: m_CamZ = m_CamZ - fwdZ * spd
        If m_Input.GetKey(Win32GL.VK_A) Then m_CamX = m_CamX - fwdZ * spd: m_CamZ = m_CamZ + fwdX * spd
        If m_Input.GetKey(Win32GL.VK_D) Then m_CamX = m_CamX + fwdZ * spd: m_CamZ = m_CamZ - fwdX * spd
        If m_Input.GetKey(&H51) Then m_CamY = m_CamY + spd  ' Q up
        If m_Input.GetKey(&H45) Then m_CamY = m_CamY - spd  ' E down

        ' Brightness
        If m_Input.GetKeyDown(&HBB) Then m_BrightScale = m_BrightScale + 0.1!  ' + key
        If m_Input.GetKeyDown(&HBD) Then m_BrightScale = m_BrightScale - 0.1!  ' - key
        If m_BrightScale < 0.1 Then m_BrightScale = 0.1!
        If m_BrightScale > 3.0 Then m_BrightScale = 3.0!

        ' Colour mode cycle
        If m_Input.GetKeyDown(&H43) Then  ' C
            m_ColorMode = (m_ColorMode + 1) Mod 3
            RebuildStarVBO
            Debug.Print "[StarMap] Color mode: " & Array("B-V Index", "Spectral Class", "White")(m_ColorMode)
        End If

        ' Highlight nearest star in Excel
        If m_Input.GetKeyDown(&H48) Then  ' H
            HighlightNearestStar
        End If

        ' Mouse look
        Dim mx As Long: mx = m_Input.GetMouseX
        Dim my As Long: my = m_Input.GetMouseY
        m_Yaw   = m_Yaw   + CSng(mx - prevMouseX) * 0.2!
        m_Pitch = m_Pitch + CSng(my - prevMouseY) * 0.15!
        If m_Pitch > 89  Then m_Pitch = 89!
        If m_Pitch < -89 Then m_Pitch = -89!
        prevMouseX = mx: prevMouseY = my

        ' --- Render ---
        GL.glClearColor 0.0, 0.0, 0.02, 1!
        GL.glClear GL.GL_COLOR_BUFFER_BIT Or GL.GL_DEPTH_BUFFER_BIT

        DrawStars

        m_Running = Win32GL.PumpMessages() And m_Running
        modWindowManager.PageFlip

        modPerf.EndFrame
        modPerf.UpdateWindowTitle hWnd
    Loop

    CleanUp
    modWindowManager.CloseGLWindow
    Debug.Print "[StarMap] Exited. " & modPerf.TotalFrames & " frames, " & m_StarCount & " stars."
    modPerf.DebugPrint
End Sub

' ============================================================
' LOAD STARS FROM EXCEL AND BUILD GPU BUFFER
' ============================================================
Private Sub InitStars()
    Dim lastRow As Long
    lastRow = modExcelBridge.GetLastRow(SHEET_NAME, modExcelBridge.STAR_COL_X)
    If lastRow < DATA_START Then Exit Sub

    Dim n As Long: n = lastRow - DATA_START + 1
    If n > MAX_STARS Then n = MAX_STARS

    ' Read raw position data for CPU picking
    ReDim m_StarPosX(0 To n - 1)
    ReDim m_StarPosY(0 To n - 1)
    ReDim m_StarPosZ(0 To n - 1)
    ReDim m_StarNames(0 To n - 1)

    ' Build interleaved VBO: x,y,z,r,g,b,size
    Dim vbuf() As Single
    ReDim vbuf(0 To n * FLOATS_PER_STAR - 1)

    Dim ws As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    If ws Is Nothing Then Exit Sub

    Dim i As Long
    For i = 0 To n - 1
        Dim row As Long: row = DATA_START + i
        Dim x As Single: x = CSng(ws.Cells(row, modExcelBridge.STAR_COL_X).Value)
        Dim y As Single: y = CSng(ws.Cells(row, modExcelBridge.STAR_COL_Y).Value)
        Dim z As Single: z = CSng(ws.Cells(row, modExcelBridge.STAR_COL_Z).Value)
        Dim mag As Single: mag = CSng(ws.Cells(row, modExcelBridge.STAR_COL_MAG).Value)
        Dim ci  As Single: ci  = CSng(ws.Cells(row, modExcelBridge.STAR_COL_CI).Value)
        Dim spect As String: spect = CStr(ws.Cells(row, 7).Value)

        m_StarPosX(i) = x
        m_StarPosY(i) = y
        m_StarPosZ(i) = z
        m_StarNames(i) = CStr(ws.Cells(row, 1).Value)

        Dim r As Single, g As Single, b As Single
        Select Case m_ColorMode
            Case 0: StellarColorMap.BVtoRGB ci, r, g, b
            Case 1: StellarColorMap.SpectralClassToRGB spect, r, g, b
            Case 2: r = 1!: g = 1!: b = 1!
        End Select

        Dim sz As Single: sz = StellarColorMap.MagnitudeToSize(mag) * m_BrightScale

        Dim base As Long: base = i * FLOATS_PER_STAR
        vbuf(base)     = x
        vbuf(base + 1) = y
        vbuf(base + 2) = z
        vbuf(base + 3) = r
        vbuf(base + 4) = g
        vbuf(base + 5) = b
        vbuf(base + 6) = sz
    Next i

    m_StarCount = n
    On Error GoTo 0

    ' Upload to GPU
    GL.glGenVertexArrays 1, m_VAO
    GL.glBindVertexArray m_VAO
    GL.glGenBuffers 1, m_VBO
    GL.glBindBuffer GL.GL_ARRAY_BUFFER, m_VBO

    Dim byteSize As LongPtr
    byteSize = n * FLOATS_PER_STAR * 4

    GL.glBufferData GL.GL_ARRAY_BUFFER, byteSize, VarPtr(vbuf(0)), GL.GL_STATIC_DRAW

    Dim stride As Long: stride = FLOATS_PER_STAR * 4   ' 7 floats * 4 bytes
    GL.glEnableVertexAttribArray 0
    GL.glVertexAttribPointer 0, 3, GL.GL_FLOAT, GL.GL_FALSE, stride, 0        ' xyz
    GL.glEnableVertexAttribArray 1
    GL.glVertexAttribPointer 1, 3, GL.GL_FLOAT, GL.GL_FALSE, stride, 12       ' rgb
    GL.glEnableVertexAttribArray 2
    GL.glVertexAttribPointer 2, 1, GL.GL_FLOAT, GL.GL_FALSE, stride, 24       ' size

    GL.glBindVertexArray 0

    ' Compile shader
    Set m_Shader = New ShaderProgram
    m_Shader.CreateFromSource EmbeddedShaders.STAR_VERTEX, EmbeddedShaders.STAR_FRAGMENT

    Debug.Print "[StarMap] Loaded " & m_StarCount & " stars into GPU buffer."
End Sub

Private Sub RebuildStarVBO()
    ' Destroy existing and rebuild with new colour mode
    If m_VBO <> 0 Then GL.glDeleteBuffers 1, m_VBO
    If m_VAO <> 0 Then GL.glDeleteVertexArrays 1, m_VAO
    m_StarCount = 0
    InitStars
End Sub

' ============================================================
' DRAW STARS
' ============================================================
Private Sub DrawStars()
    If m_StarCount = 0 Or m_VAO = 0 Then Exit Sub

    m_Shader.Use

    ' Build view matrix from camera position and angles
    Dim view As GLMath.Mat4
    view = BuildViewMatrix()

    Dim proj As GLMath.Mat4
    proj = GLMath.Perspective(60, 1024 / 768, 0.1, 2000)

    m_Shader.SetUniformMat4 "view", view
    m_Shader.SetUniformMat4 "projection", proj
    m_Shader.SetUniform1f "screenH", 768

    GL.glBindVertexArray m_VAO
    modGL.glDrawArrays GL_POINTS, 0, m_StarCount
    GL.glBindVertexArray 0

    modPerf.CountDraw m_StarCount
End Sub

Private Function BuildViewMatrix() As GLMath.Mat4
    Dim yawR As Single:   yawR   = m_Yaw   * 3.14159265 / 180!
    Dim pitchR As Single: pitchR = m_Pitch * 3.14159265 / 180!

    Dim fwdX As Single: fwdX = CSng(Cos(pitchR) * Sin(yawR))
    Dim fwdY As Single: fwdY = CSng(Sin(pitchR))
    Dim fwdZ As Single: fwdZ = CSng(-Cos(pitchR) * Cos(yawR))

    BuildViewMatrix = GLMath.LookAt( _
        m_CamX, m_CamY, m_CamZ, _
        m_CamX + fwdX, m_CamY + fwdY, m_CamZ + fwdZ, _
        0, 1, 0)
End Function

' ============================================================
' HIGHLIGHT NEAREST STAR IN EXCEL
' ============================================================
Private Sub HighlightNearestStar()
    If m_StarCount = 0 Then Exit Sub

    Dim bestDist As Single: bestDist = 1E+30
    Dim bestIdx  As Long:   bestIdx  = -1
    Dim i As Long

    For i = 0 To m_StarCount - 1
        Dim dx As Single: dx = m_StarPosX(i) - m_CamX
        Dim dy As Single: dy = m_StarPosY(i) - m_CamY
        Dim dz As Single: dz = m_StarPosZ(i) - m_CamZ
        Dim d  As Single: d  = dx * dx + dy * dy + dz * dz
        If d < bestDist Then
            bestDist = d
            bestIdx  = i
        End If
    Next i

    If bestIdx >= 0 Then
        m_HighlightRow = DATA_START + bestIdx
        modExcelBridge.HighlightRow SHEET_NAME, m_HighlightRow
        Debug.Print "[StarMap] Nearest star: " & m_StarNames(bestIdx) & _
                    " at dist=" & Format(Sqr(bestDist), "0.0") & " pc"
    End If
End Sub

' ============================================================
' CLEANUP
' ============================================================
Private Sub CleanUp()
    If Not m_Shader Is Nothing Then m_Shader.Destroy
    If m_VBO <> 0 Then GL.glDeleteBuffers 1, m_VBO:      m_VBO = 0
    If m_VAO <> 0 Then GL.glDeleteVertexArrays 1, m_VAO: m_VAO = 0
End Sub
