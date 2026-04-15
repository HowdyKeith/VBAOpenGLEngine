Option Explicit

' ============================================================
' DemoSpectra.bas  v1.0  WEEK 3
' 3D Emission / Absorption Spectra Visualiser from Excel data.
'
' SETUP:
'   1. Run modExcelBridge.GenerateSampleSpectraData to create the
'      "Spectra" sheet with hydrogen and oxygen emission lines, OR
'      paste your own data:
'        Col A = Wavelength (nm)
'        Col B = Series 1 intensity (e.g. H-alpha)
'        Col C = Series 2 intensity (e.g. OIII)
'        Col D = Series 3 (custom)
'   2. Call DemoSpectra.Run
'
' CONTROLS:
'   1/2/3     - show series 1/2/3 or overlay all
'   G         - toggle glow effect
'   R         - rotate / freeze
'   Mouse     - orbit camera
'   ESC       - quit
' ============================================================

Private Const SHEET_NAME  As String = "Spectra"
Private Const DATA_START  As Long   = 2
Private Const MAX_SERIES  As Long   = 3

' GPU resources
Private m_VAOs()    As Long   ' one VAO per series
Private m_VBOs()    As Long
Private m_Counts()  As Long   ' vertex count per series
Private m_Shader    As ShaderProgram

' State
Private m_ShowSeries As Boolean   ' True = all overlaid
Private m_ActiveSeries As Long    ' 1-3
Private m_GlowOn     As Boolean
Private m_Rotating   As Boolean
Private m_Angle      As Single
Private m_CamDist    As Single
Private m_CamPitch   As Single
Private m_Input      As InputSystem
Private m_MaxHeight  As Single

' ============================================================
' ENTRY POINT
' ============================================================
Public Sub Run()
    If Not modExcelBridge.SheetExists(SHEET_NAME) Then
        Debug.Print "[Spectra] Creating sample spectra data..."
        modExcelBridge.GenerateSampleSpectraData SHEET_NAME
    End If

    Dim hWnd As LongPtr
    hWnd = modWindowManager.CreateGLWindow("VBA Spectra Viewer  |  " & SHEET_NAME, 1024, 600)
    If hWnd = 0 Then MsgBox "Failed to create GL window", vbCritical: Exit Sub

    Dim hLib As LongPtr
    hLib = Win32GL.GetModuleHandle("opengl32.dll")
    GLLoader.Init hLib: GL.GL_Init: modPerf.Init

    GL.glEnable GL.GL_DEPTH_TEST
    GL.glEnable GL.GL_BLEND
    GL.glBlendFunc GL.GL_SRC_ALPHA, GL.GL_ONE_MINUS_SRC_ALPHA

    ' Init arrays
    ReDim m_VAOs(0 To MAX_SERIES - 1)
    ReDim m_VBOs(0 To MAX_SERIES - 1)
    ReDim m_Counts(0 To MAX_SERIES - 1)

    m_ShowSeries   = True
    m_ActiveSeries = 0    ' 0 = all
    m_GlowOn       = True
    m_Rotating     = True
    m_Angle        = 0
    m_CamDist      = 8!
    m_CamPitch     = 25!

    BuildAllBars
    If m_Counts(0) = 0 Then
        MsgBox "No spectra data on sheet '" & SHEET_NAME & "'", vbExclamation
        modWindowManager.CloseGLWindow: Exit Sub
    End If

    Set m_Input = New InputSystem
    m_Input.SetWindowHandle hWnd

    Dim t0 As Double: t0 = Win32GL.GetTime()
    Dim running As Boolean: running = True

    Do While running
        modPerf.BeginFrame

        Dim t1 As Double: t1 = Win32GL.GetTime()
        Dim dt As Single: dt = CSng(t1 - t0): t0 = t1
        If dt > 0.05 Then dt = 0.05

        m_Input.Update

        If m_Input.GetKey(Win32GL.VK_ESCAPE) Then running = False

        If m_Input.GetKeyDown(&H31) Then m_ActiveSeries = 1   ' 1
        If m_Input.GetKeyDown(&H32) Then m_ActiveSeries = 2   ' 2
        If m_Input.GetKeyDown(&H33) Then m_ActiveSeries = 3   ' 3
        If m_Input.GetKeyDown(&H30) Then m_ActiveSeries = 0   ' 0 = all
        If m_Input.GetKeyDown(&H47) Then m_GlowOn   = Not m_GlowOn    ' G
        If m_Input.GetKeyDown(&H52) Then m_Rotating = Not m_Rotating  ' R

        If m_Rotating Then m_Angle = m_Angle + 20 * dt
        If m_Angle >= 360 Then m_Angle = m_Angle - 360

        ' Mouse orbit
        m_CamPitch = m_CamPitch + CSng(m_Input.GetMouseDeltaY()) * 0.3!
        If m_CamPitch < 5  Then m_CamPitch = 5!
        If m_CamPitch > 80 Then m_CamPitch = 80!

        DrawFrame

        running = Win32GL.PumpMessages() And running
        modWindowManager.PageFlip

        modPerf.EndFrame
        modPerf.UpdateWindowTitle hWnd
    Loop

    CleanUp
    modWindowManager.CloseGLWindow
End Sub

' ============================================================
' BUILD BAR GEOMETRY FROM EXCEL DATA
' Each bar = 2 triangles (quad). Layout: x,y,z, r,g,b  (6 floats/vertex)
' ============================================================
Private Sub BuildAllBars()
    Dim ws As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = modExcelBridge.GetLastRow(SHEET_NAME, 1)
    If lastRow < DATA_START Then Exit Sub

    Dim n As Long: n = lastRow - DATA_START + 1
    m_MaxHeight = 0!

    Set m_Shader = New ShaderProgram
    m_Shader.CreateFromSource EmbeddedShaders.SPECTRA_VERTEX, EmbeddedShaders.SPECTRA_FRAGMENT

    Dim s As Long
    For s = 0 To MAX_SERIES - 1
        Dim col As Long: col = 2 + s   ' columns B, C, D
        ' 6 vertices per bar (2 triangles), 6 floats each
        Dim vbuf() As Single
        ReDim vbuf(0 To n * 6 * 6 - 1)
        Dim vi As Long: vi = 0

        Dim i As Long
        For i = 0 To n - 1
            Dim row As Long: row = DATA_START + i
            Dim wl  As Single: wl  = CSng(ws.Cells(row, 1).Value)
            Dim ht  As Single: ht  = CSng(ws.Cells(row, col).Value)
            If IsNumeric(ws.Cells(row, col).Value) Then
                If ht < 0 Then ht = 0!
            Else
                ht = 0!
            End If
            If ht > m_MaxHeight Then m_MaxHeight = ht

            ' Map wavelength 380-700nm to x position -5..+5
            Dim bx As Single: bx = CSng((wl - 380) / (700 - 380) * 10 - 5)
            Dim bw As Single: bw = 0.08!   ' bar width

            ' Colour = wavelength RGB
            Dim r As Single, g As Single, b As Single
            StellarColorMap.WavelengthToRGB wl, r, g, b

            ' Two triangles forming a quad (bottom-left origin)
            ' Bottom-left, Bottom-right, Top-right, Bottom-left, Top-right, Top-left
            Dim verts(0 To 35) As Single
            ' v0 bottom-left
            verts(0) = bx:       verts(1) = 0:  verts(2) = 0: verts(3) = r:  verts(4) = g:  verts(5) = b
            ' v1 bottom-right
            verts(6) = bx + bw:  verts(7) = 0:  verts(8) = 0: verts(9) = r:  verts(10) = g: verts(11) = b
            ' v2 top-right
            verts(12) = bx + bw: verts(13) = ht: verts(14) = 0: verts(15) = r: verts(16) = g: verts(17) = b
            ' v3 bottom-left (repeat)
            verts(18) = bx:      verts(19) = 0:  verts(20) = 0: verts(21) = r: verts(22) = g: verts(23) = b
            ' v4 top-right (repeat)
            verts(24) = bx + bw: verts(25) = ht: verts(26) = 0: verts(27) = r: verts(28) = g: verts(29) = b
            ' v5 top-left
            verts(30) = bx:      verts(31) = ht: verts(32) = 0: verts(33) = r: verts(34) = g: verts(35) = b

            Dim j As Long
            For j = 0 To 35
                vbuf(vi) = verts(j): vi = vi + 1
            Next j
        Next i

        m_Counts(s) = n * 6

        GL.glGenVertexArrays 1, m_VAOs(s)
        GL.glBindVertexArray m_VAOs(s)
        GL.glGenBuffers 1, m_VBOs(s)
        GL.glBindBuffer GL.GL_ARRAY_BUFFER, m_VBOs(s)
        GL.glBufferData GL.GL_ARRAY_BUFFER, n * 6 * 6 * 4, VarPtr(vbuf(0)), GL.GL_STATIC_DRAW

        Dim stride As Long: stride = 6 * 4
        GL.glEnableVertexAttribArray 0
        GL.glVertexAttribPointer 0, 3, GL.GL_FLOAT, GL.GL_FALSE, stride, 0
        GL.glEnableVertexAttribArray 1
        GL.glVertexAttribPointer 1, 3, GL.GL_FLOAT, GL.GL_FALSE, stride, 12
        GL.glBindVertexArray 0

        Debug.Print "[Spectra] Series " & (s + 1) & ": " & n & " bars loaded."
    Next s
    On Error GoTo 0
End Sub

' ============================================================
' DRAW
' ============================================================
Private Sub DrawFrame()
    GL.glClearColor 0.05, 0.05, 0.08, 1!
    GL.glClear GL.GL_COLOR_BUFFER_BIT Or GL.GL_DEPTH_BUFFER_BIT

    m_Shader.Use

    ' Camera orbiting the spectra
    Dim pitchR As Single: pitchR = m_CamPitch * 3.14159 / 180
    Dim angleR As Single: angleR = m_Angle    * 3.14159 / 180
    Dim eyeX As Single:   eyeX   = CSng(m_CamDist * Cos(pitchR) * Sin(angleR))
    Dim eyeY As Single:   eyeY   = CSng(m_CamDist * Sin(pitchR))
    Dim eyeZ As Single:   eyeZ   = CSng(m_CamDist * Cos(pitchR) * Cos(angleR))

    Dim view As GLMath.Mat4
    view = GLMath.LookAt(eyeX, eyeY, eyeZ, 0, 0.5, 0, 0, 1, 0)

    Dim proj As GLMath.Mat4
    proj = GLMath.Perspective(55, 1024 / 600, 0.1, 200)

    ' Identity model matrix
    Dim model As GLMath.Mat4
    modMathUtils.IdentityMatrix model

    m_Shader.SetUniformMat4 "model", model
    m_Shader.SetUniformMat4 "view", view
    m_Shader.SetUniformMat4 "projection", proj
    m_Shader.SetUniform1f "maxHeight", m_MaxHeight
    m_Shader.SetUniform1f "glowStrength", IIf(m_GlowOn, 1.5!, 0!)

    Dim s As Long
    For s = 0 To MAX_SERIES - 1
        If m_ActiveSeries = 0 Or m_ActiveSeries = s + 1 Then
            If m_Counts(s) > 0 Then
                GL.glBindVertexArray m_VAOs(s)
                modGL.glDrawArrays GL_TRIANGLES, 0, m_Counts(s)
                GL.glBindVertexArray 0
                modPerf.CountDraw m_Counts(s) \ 3
            End If
        End If
    Next s
End Sub

Private Sub CleanUp()
    If Not m_Shader Is Nothing Then m_Shader.Destroy
    Dim s As Long
    For s = 0 To MAX_SERIES - 1
        If m_VBOs(s) <> 0 Then GL.glDeleteBuffers 1, m_VBOs(s)
        If m_VAOs(s) <> 0 Then GL.glDeleteVertexArrays 1, m_VAOs(s)
    Next s
End Sub
