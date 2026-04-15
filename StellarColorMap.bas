Option Explicit

' ============================================================
' StellarColorMap.bas  v1.0  WEEK 3
' Converts stellar/spectral data to RGB colours.
'   - B-V color index  -> RGB (star surface temperature)
'   - Spectral class   -> RGB (O B A F G K M)
'   - Wavelength (nm)  -> RGB (CIE 1931 visible spectrum)
'   - Temperature (K)  -> RGB (blackbody approximation)
' All colours returned as Single in 0..1 range.
' ============================================================

' ============================================================
' B-V COLOR INDEX -> RGB
' B-V range: -0.4 (hottest blue O-stars) to +2.0 (coolest red M-stars)
' Approximation by Ballesteros (2012) via temperature then blackbody.
' ============================================================
Public Sub BVtoRGB(ByVal bv As Single, _
                   ByRef r As Single, ByRef g As Single, ByRef b As Single)
    ' Clamp B-V to valid range
    If bv < -0.4 Then bv = -0.4!
    If bv > 2.0  Then bv = 2.0!

    ' Convert B-V to temperature (Ballesteros 2012 approximation)
    Dim temp As Single
    temp = CSng(4600 * (1 / (0.92 * bv + 1.7) + 1 / (0.92 * bv + 0.62)))

    TemperatureToRGB temp, r, g, b
End Sub

' ============================================================
' BLACKBODY TEMPERATURE -> RGB (Planckian locus approximation)
' Valid range roughly 1000K - 40000K.
' Based on Tanner Helland's algorithm with stellar extension.
' ============================================================
Public Sub TemperatureToRGB(ByVal tempK As Single, _
                             ByRef r As Single, ByRef g As Single, ByRef b As Single)
    Dim t As Single: t = tempK / 100!

    ' --- RED ---
    If t <= 66 Then
        r = 1!
    Else
        Dim rv As Single: rv = CSng(329.698727446 * ((t - 60) ^ -0.1332047592))
        r = CSng(rv / 255!)
        If r < 0 Then r = 0!
        If r > 1 Then r = 1!
    End If

    ' --- GREEN ---
    Dim gv As Single
    If t <= 66 Then
        gv = CSng(99.4708025861 * Log(t) - 161.1195681661)
    Else
        gv = CSng(288.1221695283 * ((t - 60) ^ -0.0755148492))
    End If
    g = CSng(gv / 255!)
    If g < 0 Then g = 0!
    If g > 1 Then g = 1!

    ' --- BLUE ---
    If t >= 66 Then
        b = 1!
    ElseIf t <= 19 Then
        b = 0!
    Else
        Dim bv As Single: bv = CSng(138.5177312231 * Log(t - 10) - 305.0447927307)
        b = CSng(bv / 255!)
        If b < 0 Then b = 0!
        If b > 1 Then b = 1!
    End If
End Sub

' ============================================================
' SPECTRAL CLASS -> RGB (classic OBAFGKM mnemonic)
' Returns the characteristic colour of each spectral class.
' ============================================================
Public Sub SpectralClassToRGB(ByVal spect As String, _
                               ByRef r As Single, ByRef g As Single, ByRef b As Single)
    Dim cls As String: cls = UCase(Left$(Trim$(spect), 1))

    Select Case cls
        Case "O": r = 0.61: g = 0.71: b = 1.00   ' hot blue-white
        Case "B": r = 0.73: g = 0.84: b = 1.00   ' blue-white
        Case "A": r = 0.96: g = 0.96: b = 1.00   ' white
        Case "F": r = 1.00: g = 1.00: b = 0.87   ' yellow-white
        Case "G": r = 1.00: g = 0.95: b = 0.70   ' yellow (Sun-like)
        Case "K": r = 1.00: g = 0.80: b = 0.50   ' orange
        Case "M": r = 1.00: g = 0.55: b = 0.35   ' red-orange
        Case "L": r = 0.90: g = 0.35: b = 0.20   ' dark red (brown dwarf)
        Case "T": r = 0.70: g = 0.25: b = 0.10   ' methane dwarf
        Case "W": r = 0.40: g = 0.60: b = 1.00   ' Wolf-Rayet (blue)
        Case Else: r = 1!:  g = 1!:   b = 1!     ' unknown = white
    End Select
End Sub

' ============================================================
' WAVELENGTH (nm) -> RGB  (CIE 1931 visible spectrum)
' Range: 380nm (violet) - 700nm (red). Outside = black.
' Based on Dan Bruton's algorithm (physics.sfasu.edu).
' ============================================================
Public Sub WavelengthToRGB(ByVal wl As Single, _
                            ByRef r As Single, ByRef g As Single, ByRef b As Single)
    r = 0!: g = 0!: b = 0!

    If wl < 380 Or wl > 700 Then Exit Sub

    If wl < 440 Then
        r = CSng(-(wl - 440) / (440 - 380))
        b = 1!
    ElseIf wl < 490 Then
        g = CSng((wl - 440) / (490 - 440))
        b = 1!
    ElseIf wl < 510 Then
        g = 1!
        b = CSng(-(wl - 510) / (510 - 490))
    ElseIf wl < 580 Then
        r = CSng((wl - 510) / (580 - 510))
        g = 1!
    ElseIf wl < 645 Then
        r = 1!
        g = CSng(-(wl - 645) / (645 - 580))
    Else
        r = 1!
    End If

    ' Intensity falloff at edges of visible range
    Dim factor As Single
    If wl < 420 Then
        factor = CSng(0.3 + 0.7 * (wl - 380) / (420 - 380))
    ElseIf wl > 680 Then
        factor = CSng(0.3 + 0.7 * (700 - wl) / (700 - 680))
    Else
        factor = 1!
    End If

    r = CSng(r * factor)
    g = CSng(g * factor)
    b = CSng(b * factor)
End Sub

' ============================================================
' DENSITY -> COLOUR (for gas/volume rendering)
' Maps a 0..1 density to a colour using a chosen palette.
' ============================================================
Public Sub DensityToRGB(ByVal density As Single, _
                         ByRef r As Single, ByRef g As Single, ByRef b As Single, _
                         Optional ByVal palette As Long = 0)
    ' Clamp
    If density < 0 Then density = 0!
    If density > 1 Then density = 1!

    Select Case palette
        Case 0   ' Plasma (purple -> orange -> yellow)
            r = CSng(0.05 + density * 0.95)
            g = CSng(density * density * 0.8)
            b = CSng(0.5 * (1 - density))

        Case 1   ' Inferno (black -> red -> yellow -> white)
            If density < 0.5 Then
                r = CSng(density * 2)
                g = 0!
                b = 0!
            Else
                r = 1!
                g = CSng((density - 0.5) * 2)
                b = CSng((density - 0.5) * 2 * 0.5)
            End If

        Case 2   ' Viridis (dark blue -> teal -> yellow)
            r = CSng(density * 0.9)
            g = CSng(0.4 + density * 0.6)
            b = CSng(0.5 * (1 - density) + 0.1)

        Case 3   ' Oxygen nebula (blue-green dominant)
            r = CSng(density * density * 0.3)
            g = CSng(0.6 * density)
            b = CSng(0.8 * density)

        Case Else
            r = density: g = density: b = density
    End Select
End Sub

' ============================================================
' MAGNITUDE -> VISUAL SIZE (for point sprite star rendering)
' Returns a 0..1 size suitable for gl_PointSize scaling.
' Apparent mag range: -1.5 (Sirius) to +6.5 (naked-eye limit)
' ============================================================
Public Function MagnitudeToSize(ByVal mag As Single) As Single
    ' Inverse: brighter (lower mag) = larger point
    Dim size As Single
    size = CSng(1.2 - mag * 0.12)   ' linear approximation
    If size < 0.05 Then size = 0.05!
    If size > 1.5  Then size = 1.5!
    MagnitudeToSize = size
End Function
