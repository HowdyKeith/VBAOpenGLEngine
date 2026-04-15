Option Explicit

' ============================================================
' DemoGasDensity.bas  v1.0  WEEK 3
' 3D Gas / Oxygen Density Volume Visualiser from Excel data.
' Uses additive-blended billboard quads sorted back-to-front
' for a soft volumetric appearance.
'
' SETUP:
'   1. Run modExcelBridge.GenerateSampleGasDensity to create the
'      "GasDensity" sheet, OR provide your own data:
'        Col A=X, B=Y, C=Z, D=Density(0..1), E=R, F=G, G=B
'   2. Call DemoGasDensity.Run
'
' CONTROLS:
'   Mouse drag    - orbit camera
'   W/S           - zoom in/out
'   P             - cycle colour palette
'   T             - toggle transparency mode (additive / alpha-blend)
'   ESC           - quit
' ============================================================

Private Const SHEET_NAME  As String = "GasDensity"
Private Const DATA_START  As Long   = 2
Private Const MAX_POINTS  As Long   = 5000

' Interleaved layout per quad vertex: x,y,z, u,v, r,g,b,a  (9 floats)
Private Const FLOATS_PER_VERT As Long = 9
Private Const VERTS_PER_QUAD  As Long = 6   ' 2 triangles

Private m_VAO       As Long
Private m_VBO       As Long
Private m_Shader    As ShaderProgram
Private m_Count     As Long   ' number of quads

' CPU-side position for sorting
Private m_PosX()  As Single
Private m_PosY()  As Single
Private m_PosZ()  As Single
Private m_SortIdx() As Long

' State
Private m_CamDist  As Single
Private m_Yaw      As Single
Private m_Pitch    As Single
Private m_Palette  As Long
Private m_Additive As Boolean
Private m_Input    As InputSystem

' ============================================================
' ENTRY POINT
' ============================================================
Public Sub Run()
    If Not modExcelBridge.SheetExists(SHEET_NAME) Then
        Debug.Print "[GasDensity] Creating sample data..."
        modExcelBridge.GenerateSampleGasDensity SHEET_NAME, 1500
    End If

    Dim hWnd As LongPtr
    hWnd = modWindowManager.CreateGLWindow("VBA Gas Density  |  " & SHEET_NAME, 900, 700)
    If hWnd = 0 Then MsgBox "Failed to create GL window", vbCritical: Exit Sub

    Dim hLib As LongPtr
    hLib = Win32GL.GetModuleHandle("opengl32.dll")
    GLLoader.Init hLib: GL.GL_Init: modPerf.Init

    GL.glEnable GL.GL_BLEND
    GL.glDisable GL.GL_DEPTH_TEST   ' depth writes OFF for correct volume blending
    GL.glBlendFunc GL.GL_SRC_ALPHA, GL.GL_ONE  ' additive by default

    m_CamDist = 20!
    m_Yaw     = 0!
    m_Pitch   = 25!
    m_Palette  = 3   ' Oxygen palette
    m_Additive = True

    LoadData
    If m_Count = 0 Then
        MsgBox "No density data on sheet '" & SHEET_NAME & "'", vbExclamation
        modWindowManager.CloseGLWindow: Exit Sub
    End If

    Set m_Input = New InputSystem
    m_Input.SetWindowHandle hWnd

    Dim t0 As Double: t0 = Win32GL.GetTime()
    Dim running As Boolean: running = True
    Dim prevMX As Long: prevMX = m_Input.GetMouseX
    Dim prevMY As Long: prevMY = m_Input.GetMouseY

    Do While running
        modPerf.BeginFrame

        Dim t1 As Double: t1 = Win32GL.GetTime()
        Dim dt As Single: dt = CSng(t1 - t0): t0 = t1
        If dt > 0.05 Then dt = 0.05

        m_Input.Update

        If m_Input.GetKey(Win32GL.VK_ESCAPE) Then running = False

        ' Mouse orbit
        Dim mx As Long: mx = m_Input.GetMouseX
        Dim my As Long: my = m_Input.GetMouseY
        m_Yaw   = m_Yaw   + CSng(mx - prevMX) * 0.4!
        m_Pitch = m_Pitch + CSng(my - prevMY) * 0.3!
        If m_Pitch < 5  Then m_Pitch = 5!
        If m_Pitch > 85 Then m_Pitch = 85!
        prevMX = mx: prevMY = my

        If m_Input.GetKey(Win32GL.VK_W) Then m_CamDist = m_CamDist - 10 * dt
        If m_Input.GetKey(Win32GL.VK_S) Then m_CamDist = m_CamDist + 10 * dt
        If m_CamDist < 2 Then m_CamDist = 2!
        If m_CamDist > 80 Then m_CamDist = 80!

        If m_Input.GetKeyDown(&H50) Then    ' P - palette cycle
            m_Palette = (m_Palette + 1) Mod 4
            RebuildVBO
        End If

        If m_Input.GetKeyDown(&H54) Then    ' T - toggle blend mode
            m_Additive = Not m_Additive
            If m_Additive Then
                GL.glBlendFunc GL.GL_SRC_ALPHA, GL.GL_ONE
            Else
                GL.glBlendFunc GL.GL_SRC_ALPHA, GL.GL_ONE_MINUS_SRC_ALPHA
            End If
        End If

        SortAndDraw

        running = Win32GL.PumpMessages() And running
        modWindowManager.PageFlip

        modPerf.EndFrame
        modPerf.UpdateWindowTitle hWnd
    Loop

    CleanUp
    modWindowManager.CloseGLWindow
End Sub

' ============================================================
' LOAD DATA FROM EXCEL
' ============================================================
Private Sub LoadData()
    Dim ws As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    If ws Is Nothing Then Exit Sub

    Dim lastRow As Long
    lastRow = modExcelBridge.GetLastRow(SHEET_NAME, 1)
    If lastRow < DATA_START Then Exit Sub

    Dim n As Long: n = lastRow - DATA_START + 1
    If n > MAX_POINTS Then n = MAX_POINTS

    ReDim m_PosX(0 To n - 1)
    ReDim m_PosY(0 To n - 1)
    ReDim m_PosZ(0 To n - 1)
    ReDim m_SortIdx(0 To n - 1)

    ' Store raw positions for back-to-front sorting
    Dim i As Long
    For i = 0 To n - 1
        Dim row As Long: row = DATA_START + i
        m_PosX(i) = CSng(ws.Cells(row, 1).Value)
        m_PosY(i) = CSng(ws.Cells(row, 2).Value)
        m_PosZ(i) = CSng(ws.Cells(row, 3).Value)
        m_SortIdx(i) = i
    Next i

    m_Count = n
    On Error GoTo 0

    ' Build initial VBO
    Set m_Shader = New ShaderProgram
    m_Shader.CreateFromSource EmbeddedShaders.VOLUME_VERTEX, EmbeddedShaders.VOLUME_FRAGMENT
    RebuildVBO
End Sub

Private Sub RebuildVBO()
    Dim ws As Object
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_NAME)
    If ws Is Nothing Then Exit Sub

    ' 6 vertices per quad, 9 floats each
    Dim vbuf() As Single
    ReDim vbuf(0 To m_Count * VERTS_PER_QUAD * FLOATS_PER_VERT - 1)

    Dim i As Long, vi As Long: vi = 0
    For i = 0 To m_Count - 1
        Dim row As Long: row = DATA_START + i
        Dim x As Single: x = m_PosX(i)
        Dim y As Single: y = m_PosY(i)
        Dim z As Single: z = m_PosZ(i)

        Dim dens As Single: dens = CSng(ws.Cells(row, 4).Value)
        If dens < 0 Then dens = 0!: If dens > 1 Then dens = 1!

        Dim r As Single, g As Single, b As Single

        ' Use per-point colour if columns E,F,G are filled
        Dim hasColor As Boolean
        hasColor = IsNumeric(ws.Cells(row, 5).Value)
        If hasColor Then
            r = CSng(ws.Cells(row, 5).Value)
            g = CSng(ws.Cells(row, 6).Value)
            b = CSng(ws.Cells(row, 7).Value)
        Else
            StellarColorMap.DensityToRGB dens, r, g, b, m_Palette
        End If

        Dim alpha As Single: alpha = dens * 0.35!   ' keep transparent

        ' Billboard half-size = density-weighted
        Dim hs As Single: hs = 0.3! + dens * 0.7!

        ' Flat camera-facing quad (simplified: always face +Z for now)
        ' Full billboarding requires camera right/up vectors - simplified here
        Dim x0 As Single: x0 = x - hs
        Dim x1 As Single: x1 = x + hs
        Dim y0 As Single: y0 = y - hs
        Dim y1 As Single: y1 = y + hs

        ' 6 vertices: BL, BR, TR, BL, TR, TL
        Dim vdata(0 To 53) As Single
        ' BL
        vdata(0)=x0: vdata(1)=y0: vdata(2)=z: vdata(3)=0: vdata(4)=0: vdata(5)=r: vdata(6)=g: vdata(7)=b: vdata(8)=alpha
        ' BR
        vdata(9)=x1: vdata(10)=y0: vdata(11)=z: vdata(12)=1: vdata(13)=0: vdata(14)=r: vdata(15)=g: vdata(16)=b: vdata(17)=alpha
        ' TR
        vdata(18)=x1: vdata(19)=y1: vdata(20)=z: vdata(21)=1: vdata(22)=1: vdata(23)=r: vdata(24)=g: vdata(25)=b: vdata(26)=alpha
        ' BL
        vdata(27)=x0: vdata(28)=y0: vdata(29)=z: vdata(30)=0: vdata(31)=0: vdata(32)=r: vdata(33)=g: vdata(34)=b: vdata(35)=alpha
        ' TR
        vdata(36)=x1: vdata(37)=y1: vdata(38)=z: vdata(39)=1: vdata(40)=1: vdata(41)=r: vdata(42)=g: vdata(43)=b: vdata(44)=alpha
        ' TL
        vdata(45)=x0: vdata(46)=y1: vdata(47)=z: vdata(48)=0: vdata(49)=1: vdata(50)=r: vdata(51)=g: vdata(52)=b: vdata(53)=alpha

        Dim j As Long
        For j = 0 To 53: vbuf(vi) = vdata(j): vi = vi + 1: Next j
    Next i

    ' Upload or update GPU buffer
    If m_VAO = 0 Then
        GL.glGenVertexArrays 1, m_VAO
        GL.glBindVertexArray m_VAO
        GL.glGenBuffers 1, m_VBO
        GL.glBindBuffer GL.GL_ARRAY_BUFFER, m_VBO
        GL.glBufferData GL.GL_ARRAY_BUFFER, m_Count * VERTS_PER_QUAD * FLOATS_PER_VERT * 4, _
                        VarPtr(vbuf(0)), GL.GL_DYNAMIC_DRAW

        Dim stride As Long: stride = FLOATS_PER_VERT * 4
        GL.glEnableVertexAttribArray 0: GL.glVertexAttribPointer 0, 3, GL.GL_FLOAT, GL.GL_FALSE, stride, 0   ' xyz
        GL.glEnableVertexAttribArray 1: GL.glVertexAttribPointer 1, 2, GL.GL_FLOAT, GL.GL_FALSE, stride, 12  ' uv
        GL.glEnableVertexAttribArray 2: GL.glVertexAttribPointer 2, 4, GL.GL_FLOAT, GL.GL_FALSE, stride, 20  ' rgba
        GL.glBindVertexArray 0
    Else
        GL.glBindBuffer GL.GL_ARRAY_BUFFER, m_VBO
        GL.glBufferSubData GL.GL_ARRAY_BUFFER, 0, _
                           m_Count * VERTS_PER_QUAD * FLOATS_PER_VERT * 4, VarPtr(vbuf(0))
    End If
    On Error GoTo 0
    Debug.Print "[GasDensity] VBO built: " & m_Count & " quads, palette=" & m_Palette
End Sub

' ============================================================
' SORT QUADS BACK-TO-FRONT AND DRAW
' ============================================================
Private Sub SortAndDraw()
    GL.glClearColor 0.0, 0.0, 0.02, 1!
    GL.glClear GL.GL_COLOR_BUFFER_BIT Or GL.GL_DEPTH_BUFFER_BIT

    ' Camera position
    Dim yR As Single: yR = m_Yaw   * 3.14159265 / 180!
    Dim pR As Single: pR = m_Pitch * 3.14159265 / 180!
    Dim eyeX As Single: eyeX = CSng(m_CamDist * Cos(pR) * Sin(yR))
    Dim eyeY As Single: eyeY = CSng(m_CamDist * Sin(pR))
    Dim eyeZ As Single: eyeZ = CSng(m_CamDist * Cos(pR) * Cos(yR))

    ' Simple insertion sort by distance to camera (adequate for ~1000 quads)
    ' For production use a radix sort or GPU sort
    Dim i As Long, j As Long
    For i = 1 To m_Count - 1
        Dim idxI As Long: idxI = m_SortIdx(i)
        Dim dxi As Single: dxi = m_PosX(idxI) - eyeX
        Dim dyi As Single: dyi = m_PosY(idxI) - eyeY
        Dim dzi As Single: dzi = m_PosZ(idxI) - eyeZ
        Dim di  As Single: di  = dxi * dxi + dyi * dyi + dzi * dzi

        j = i - 1
        Do While j >= 0
            Dim idxJ As Long: idxJ = m_SortIdx(j)
            Dim dxj As Single: dxj = m_PosX(idxJ) - eyeX
            Dim dyj As Single: dyj = m_PosY(idxJ) - eyeY
            Dim dzj As Single: dzj = m_PosZ(idxJ) - eyeZ
            Dim dj  As Single: dj  = dxj * dxj + dyj * dyj + dzj * dzj
            If dj >= di Then Exit Do
            m_SortIdx(j + 1) = m_SortIdx(j)
            j = j - 1
        Loop
        m_SortIdx(j + 1) = idxI
    Next i

    m_Shader.Use

    Dim view As GLMath.Mat4
    view = GLMath.LookAt(eyeX, eyeY, eyeZ, 0, 0, 0, 0, 1, 0)
    Dim proj As GLMath.Mat4
    proj = GLMath.Perspective(55, 900! / 700!, 0.1, 200)

    m_Shader.SetUniformMat4 "view", view
    m_Shader.SetUniformMat4 "projection", proj

    GL.glBindVertexArray m_VAO
    ' Draw sorted back-to-front quad by quad
    For i = m_Count - 1 To 0 Step -1
        Dim idx As Long: idx = m_SortIdx(i)
        modGL.glDrawArrays GL_TRIANGLES, idx * VERTS_PER_QUAD, VERTS_PER_QUAD
    Next i
    GL.glBindVertexArray 0

    modPerf.CountDraw m_Count * 2
End Sub

Private Sub CleanUp()
    If Not m_Shader Is Nothing Then m_Shader.Destroy
    If m_VBO <> 0 Then GL.glDeleteBuffers 1, m_VBO
    If m_VAO <> 0 Then GL.glDeleteVertexArrays 1, m_VAO
End Sub
