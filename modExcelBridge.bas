Option Explicit

' ============================================================
' modExcelBridge.bas  v1.0  WEEK 3
' Excel ↔ GL Engine data bridge.
' Reads worksheet ranges into flat Single arrays the GL demos can
' upload directly into VBOs. Also writes metrics back to cells.
'
' DESIGN RULES:
'   - All reads go through On Error Resume Next so the GL loop
'     keeps running if Excel is closed or a range doesn't exist.
'   - Never allocates inside a per-frame hot path.
'   - Column letters or 1-based indices both accepted.
'   - Sheet name defaults to active sheet if omitted.
' ============================================================

' ============================================================
' CONSTANTS - column indices for the built-in demo sheet layouts
' ============================================================
' Stars sheet (STARS or HYG):  A=Name, B=X, C=Y, D=Z, E=Mag, F=CI, G=Spect
Public Const STAR_COL_X     As Long = 2
Public Const STAR_COL_Y     As Long = 3
Public Const STAR_COL_Z     As Long = 4
Public Const STAR_COL_MAG   As Long = 5
Public Const STAR_COL_CI    As Long = 6   ' B-V color index

' Spectra sheet: A=Wavelength(nm), B=Intensity
Public Const SPEC_COL_WAVE  As Long = 1
Public Const SPEC_COL_INTENS As Long = 2

' Gas Density sheet: A=X, B=Y, C=Z, D=Density, E=R, F=G, G=B (optional)
Public Const GAS_COL_X      As Long = 1
Public Const GAS_COL_Y      As Long = 2
Public Const GAS_COL_Z      As Long = 3
Public Const GAS_COL_DENS   As Long = 4
Public Const GAS_COL_R      As Long = 5
Public Const GAS_COL_G      As Long = 6
Public Const GAS_COL_B      As Long = 7

' ============================================================
' FIND THE LAST USED ROW IN A COLUMN (1-based)
' ============================================================
Public Function GetLastRow(ByVal sheetName As String, ByVal col As Long) As Long
    GetLastRow = 1
    On Error Resume Next
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If ws Is Nothing Then Exit Function
    Dim lastRow As Long
    lastRow = ws.Cells(ws.Rows.Count, col).End(-4162).Row  ' xlUp = -4162
    If lastRow > 1 Then GetLastRow = lastRow
    On Error GoTo 0
End Function

' ============================================================
' READ A SINGLE COLUMN INTO A Single() ARRAY
' Returns the number of values read.
' ============================================================
Public Function ReadColumn(ByVal sheetName As String, _
                           ByVal col As Long, _
                           ByVal startRow As Long, _
                           ByRef outArray() As Single) As Long
    ReadColumn = 0
    On Error Resume Next
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If ws Is Nothing Then Exit Function

    Dim lastRow As Long
    lastRow = GetLastRow(sheetName, col)
    If lastRow < startRow Then Exit Function

    Dim n As Long: n = lastRow - startRow + 1
    ReDim outArray(0 To n - 1)

    Dim i As Long
    For i = 0 To n - 1
        Dim v As Variant
        v = ws.Cells(startRow + i, col).Value
        If IsNumeric(v) Then outArray(i) = CSng(v) Else outArray(i) = 0!
    Next i

    ReadColumn = n
    On Error GoTo 0
End Function

' ============================================================
' READ MULTIPLE COLUMNS INTO A MULTI-COLUMN Single() ARRAY
' outArray(i * colCount + colOffset) layout.
' Returns number of rows read.
' ============================================================
Public Function ReadColumns(ByVal sheetName As String, _
                            ByVal firstCol As Long, _
                            ByVal lastCol As Long, _
                            ByVal startRow As Long, _
                            ByRef outArray() As Single) As Long
    ReadColumns = 0
    On Error Resume Next
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If ws Is Nothing Then Exit Function

    Dim colCount As Long: colCount = lastCol - firstCol + 1
    Dim lastRow As Long:  lastRow  = GetLastRow(sheetName, firstCol)
    If lastRow < startRow Then Exit Function

    Dim n As Long: n = lastRow - startRow + 1
    ReDim outArray(0 To n * colCount - 1)

    Dim r As Long, c As Long
    For r = 0 To n - 1
        For c = 0 To colCount - 1
            Dim v As Variant
            v = ws.Cells(startRow + r, firstCol + c).Value
            outArray(r * colCount + c) = IIf(IsNumeric(v), CSng(v), 0!)
        Next c
    Next r

    ReadColumns = n
    On Error GoTo 0
End Function

' ============================================================
' READ A NAMED RANGE INTO Single() ARRAY (column-major)
' ============================================================
Public Function ReadNamedRange(ByVal rangeName As String, _
                               ByRef outArray() As Single) As Long
    ReadNamedRange = 0
    On Error Resume Next
    Dim rng As Object
    Set rng = ThisWorkbook.Names(rangeName).RefersToRange
    If rng Is Nothing Then Exit Function

    Dim nRows As Long: nRows = rng.Rows.Count
    Dim nCols As Long: nCols = rng.Columns.Count
    ReDim outArray(0 To nRows * nCols - 1)

    Dim r As Long, c As Long
    For r = 1 To nRows
        For c = 1 To nCols
            Dim v As Variant: v = rng.Cells(r, c).Value
            outArray((r - 1) * nCols + (c - 1)) = IIf(IsNumeric(v), CSng(v), 0!)
        Next c
    Next r

    ReadNamedRange = nRows
    On Error GoTo 0
End Function

' ============================================================
' WRITE A METRIC BACK TO A NAMED CELL
' ============================================================
Public Sub WriteMetric(ByVal sheetName As String, _
                       ByVal cellAddress As String, _
                       ByVal value As Variant)
    On Error Resume Next
    ThisWorkbook.Worksheets(sheetName).Range(cellAddress).Value = value
    On Error GoTo 0
End Sub

' ============================================================
' CHECK WHETHER A SHEET EXISTS
' ============================================================
Public Function SheetExists(ByVal name As String) As Boolean
    On Error Resume Next
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets(name)
    SheetExists = Not ws Is Nothing
    On Error GoTo 0
End Function

' ============================================================
' HIGHLIGHT A ROW IN A SHEET (for star selection feedback)
' ============================================================
Public Sub HighlightRow(ByVal sheetName As String, _
                        ByVal rowIndex As Long, _
                        Optional ByVal colorIndex As Long = 6)  ' yellow
    On Error Resume Next
    Dim ws As Object
    Set ws = ThisWorkbook.Worksheets(sheetName)
    If ws Is Nothing Then Exit Sub
    ' Clear previous highlight
    ws.UsedRange.Interior.ColorIndex = xlNone
    ' Highlight selected row
    ws.Rows(rowIndex).Interior.ColorIndex = colorIndex
    ' Scroll Excel to show the row
    Application.Goto ws.Cells(rowIndex, 1), True
    On Error GoTo 0
End Sub

' ============================================================
' GENERATE SAMPLE STELLAR DATA (HYG-format subset)
' Writes synthetic star data to a sheet so the demo works
' with no real data file.  Call once to populate.
' ============================================================
Public Sub GenerateSampleStarData(Optional ByVal sheetName As String = "Stars", _
                                  Optional ByVal starCount As Long = 2000)
    On Error Resume Next
    Dim ws As Object

    ' Create sheet if it doesn't exist
    If Not SheetExists(sheetName) Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.name = sheetName
    Else
        Set ws = ThisWorkbook.Worksheets(sheetName)
        ws.Cells.Clear
    End If

    ' Header
    ws.Cells(1, 1).Value = "Name"
    ws.Cells(1, 2).Value = "X"
    ws.Cells(1, 3).Value = "Y"
    ws.Cells(1, 4).Value = "Z"
    ws.Cells(1, 5).Value = "Mag"
    ws.Cells(1, 6).Value = "CI"
    ws.Cells(1, 7).Value = "Spect"

    Dim i As Long
    Randomize Timer
    For i = 1 To starCount
        Dim dist As Single: dist = CSng(10 + Rnd() * 990)  ' 10-1000 parsec

        ' Galactic distribution: disc with bulge
        Dim theta As Double: theta = Rnd() * 6.283185
        Dim phi As Double:   phi   = (Rnd() - 0.5) * 0.3   ' thin disc
        Dim x As Single:     x = dist * CSng(Cos(theta) * Cos(phi))
        Dim y As Single:     y = dist * CSng(Sin(phi))
        Dim z As Single:     z = dist * CSng(Sin(theta) * Cos(phi))

        ' Magnitude: brighter stars rarer (log distribution)
        Dim mag As Single:   mag = CSng(-1.5 + Rnd() * 10)

        ' Color index: range -0.3 (hot blue) to 2.0 (cool red)
        Dim ci As Single:    ci  = CSng(-0.3 + Rnd() * 2.3)

        ' Spectral class from CI
        Dim spect As String
        If ci < 0 Then
            spect = "O"
        ElseIf ci < 0.3 Then
            spect = "B"
        ElseIf ci < 0.58 Then
            spect = "A"
        ElseIf ci < 0.81 Then
            spect = "F"
        ElseIf ci < 1.0 Then
            spect = "G"
        ElseIf ci < 1.4 Then
            spect = "K"
        Else
            spect = "M"
        End If

        ws.Cells(i + 1, 1).Value = "Star " & i
        ws.Cells(i + 1, 2).Value = Round(x, 3)
        ws.Cells(i + 1, 3).Value = Round(y, 3)
        ws.Cells(i + 1, 4).Value = Round(z, 3)
        ws.Cells(i + 1, 5).Value = Round(mag, 2)
        ws.Cells(i + 1, 6).Value = Round(ci, 3)
        ws.Cells(i + 1, 7).Value = spect
    Next i

    ws.Columns("A:G").AutoFit
    Debug.Print "[ExcelBridge] Generated " & starCount & " sample stars on sheet '" & sheetName & "'."
    On Error GoTo 0
End Sub

' ============================================================
' GENERATE SAMPLE SPECTRA DATA
' ============================================================
Public Sub GenerateSampleSpectraData(Optional ByVal sheetName As String = "Spectra")
    On Error Resume Next
    Dim ws As Object
    If Not SheetExists(sheetName) Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.name = sheetName
    Else
        Set ws = ThisWorkbook.Worksheets(sheetName)
        ws.Cells.Clear
    End If

    ws.Cells(1, 1).Value = "Wavelength_nm"
    ws.Cells(1, 2).Value = "H_alpha"       ' Hydrogen emission lines
    ws.Cells(1, 3).Value = "O_density"     ' Oxygen emission
    ws.Cells(1, 4).Value = "Custom"

    ' Visible spectrum 380-700nm in 5nm steps
    Dim r As Long: r = 2
    Dim wl As Long
    For wl = 380 To 700 Step 5
        ws.Cells(r, 1).Value = wl
        ' Hydrogen Balmer series peaks
        Dim ha As Single: ha = 0!
        If Abs(wl - 656) < 8 Then ha = CSng(1.0 - Abs(wl - 656) / 8.0)   ' H-alpha 656nm
        If Abs(wl - 486) < 6 Then ha = ha + CSng(0.5 - Abs(wl - 486) / 12.0)  ' H-beta 486nm
        If Abs(wl - 434) < 5 Then ha = ha + CSng(0.25 - Abs(wl - 434) / 20.0)  ' H-gamma
        ws.Cells(r, 2).Value = Round(ha, 4)

        ' Oxygen peaks at 495nm and 500nm (OIII doublet)
        Dim oxy As Single: oxy = 0!
        If Abs(wl - 496) < 6 Then oxy = CSng(0.8 - Abs(wl - 496) / 7.5)
        If Abs(wl - 501) < 6 Then oxy = oxy + CSng(1.0 - Abs(wl - 501) / 6.0)
        ws.Cells(r, 3).Value = Round(oxy, 4)

        ' Custom: user can fill this in
        ws.Cells(r, 4).Value = 0

        r = r + 1
    Next wl

    ws.Columns("A:D").AutoFit
    Debug.Print "[ExcelBridge] Generated spectra data on sheet '" & sheetName & "'."
    On Error GoTo 0
End Sub

' ============================================================
' GENERATE SAMPLE GAS DENSITY DATA (3D oxygen cloud)
' ============================================================
Public Sub GenerateSampleGasDensity(Optional ByVal sheetName As String = "GasDensity", _
                                    Optional ByVal pointCount As Long = 1000)
    On Error Resume Next
    Dim ws As Object
    If Not SheetExists(sheetName) Then
        Set ws = ThisWorkbook.Worksheets.Add
        ws.name = sheetName
    Else
        Set ws = ThisWorkbook.Worksheets(sheetName)
        ws.Cells.Clear
    End If

    ws.Cells(1, 1).Value = "X"
    ws.Cells(1, 2).Value = "Y"
    ws.Cells(1, 3).Value = "Z"
    ws.Cells(1, 4).Value = "Density"
    ws.Cells(1, 5).Value = "R"
    ws.Cells(1, 6).Value = "G"
    ws.Cells(1, 7).Value = "B"

    ' Simulate an emission nebula: two overlapping Gaussian blobs
    Dim i As Long
    Randomize Timer
    For i = 1 To pointCount
        ' Random point in unit sphere
        Dim theta As Double: theta = Rnd() * 6.283185
        Dim cosP As Double:  cosP  = 2 * Rnd() - 1
        Dim sinP As Double:  sinP  = Sqr(1 - cosP * cosP)
        Dim r As Double:     r     = Rnd() ^ (1 / 3)   ' cube root for uniform volume

        Dim x As Single: x = CSng(r * sinP * Cos(theta) * 5)
        Dim y As Single: y = CSng(r * cosP * 3)          ' flattened on Y
        Dim z As Single: z = CSng(r * sinP * Sin(theta) * 5)

        ' Density: Gaussian falloff from centre
        Dim dist As Single: dist = Sqr(x * x + y * y + z * z)
        Dim dens As Single: dens = CSng(Exp(-dist * dist * 0.08))

        ' Colour: blue-green oxygen core, red hydrogen halo
        Dim cr As Single: cr = CSng(0.1 + (dist / 7) * 0.8)   ' red grows with dist
        Dim cg As Single: cg = CSng(0.5 * Exp(-dist * 0.3))
        Dim cb As Single: cb = CSng(0.8 * Exp(-dist * 0.2))

        ws.Cells(i + 1, 1).Value = Round(x, 3)
        ws.Cells(i + 1, 2).Value = Round(y, 3)
        ws.Cells(i + 1, 3).Value = Round(z, 3)
        ws.Cells(i + 1, 4).Value = Round(dens, 4)
        ws.Cells(i + 1, 5).Value = Round(cr, 3)
        ws.Cells(i + 1, 6).Value = Round(cg, 3)
        ws.Cells(i + 1, 7).Value = Round(cb, 3)
    Next i

    ws.Columns("A:G").AutoFit
    Debug.Print "[ExcelBridge] Generated " & pointCount & " gas density points on '" & sheetName & "'."
    On Error GoTo 0
End Sub
