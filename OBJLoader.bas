Option Explicit
' =========================================================
' OBJLoader.bas v6.3 (FIXED - normals now parsed)
' Purpose: Parse OBJ into flat interleaved vertex buffer
'          Layout per vertex: X Y Z NX NY NZ U V  (8 floats)
' =========================================================

Public Type OBJData
    vertices() As Single   ' interleaved: pos(3) + norm(3) + uv(2)
    count     As Long      ' number of vertices (not floats)
End Type

Public Function LoadOBJ(ByVal path As String) As OBJData
    Dim f As Integer: f = FreeFile
    Open path For Input As #f

    ' Source arrays
    Dim v()  As Single, vn() As Single, vt() As Single
    Dim out() As Single
    Dim vCount As Long, vnCount As Long, vtCount As Long, outCount As Long

    Do Until EOF(f)
        Dim ln As String
        Line Input #f, ln
        ln = Trim$(ln)

        Dim parts() As String

        ' -----------------------------------------------
        ' VERTEX POSITION  "v x y z"
        ' -----------------------------------------------
        If left$(ln, 2) = "v " Then
            parts = Split(ln, " ")
            vCount = vCount + 3
            ReDim Preserve v(0 To vCount - 1)
            v(vCount - 3) = CSng(parts(1))
            v(vCount - 2) = CSng(parts(2))
            v(vCount - 1) = CSng(parts(3))

        ' -----------------------------------------------
        ' VERTEX NORMAL  "vn nx ny nz"
        ' -----------------------------------------------
        ElseIf left$(ln, 3) = "vn " Then
            parts = Split(ln, " ")
            vnCount = vnCount + 3
            ReDim Preserve vn(0 To vnCount - 1)
            vn(vnCount - 3) = CSng(parts(1))
            vn(vnCount - 2) = CSng(parts(2))
            vn(vnCount - 1) = CSng(parts(3))

        ' -----------------------------------------------
        ' TEXTURE COORD  "vt u v"
        ' -----------------------------------------------
        ElseIf left$(ln, 3) = "vt " Then
            parts = Split(ln, " ")
            vtCount = vtCount + 2
            ReDim Preserve vt(0 To vtCount - 1)
            vt(vtCount - 2) = CSng(parts(1))
            vt(vtCount - 1) = 1 - CSng(parts(2))   ' flip Y for OpenGL

        ' -----------------------------------------------
        ' FACE  "f v/vt/vn ..."  (triangles only)
        ' -----------------------------------------------
        ElseIf left$(ln, 2) = "f " Then
            parts = Split(ln, " ")
            Dim i As Long
            For i = 1 To 3
                Dim tok() As String
                tok = Split(parts(i), "/")

                Dim vi As Long:  vi = CLng(tok(0)) - 1
                Dim ti As Long:  ti = -1
                Dim ni As Long:  ni = -1

                If UBound(tok) >= 1 Then
                    If Len(Trim$(tok(1))) > 0 Then ti = CLng(tok(1)) - 1
                End If
                If UBound(tok) >= 2 Then
                    If Len(Trim$(tok(2))) > 0 Then ni = CLng(tok(2)) - 1
                End If

                ' 8 floats per vertex: pos(3) norm(3) uv(2)
                outCount = outCount + 8
                ReDim Preserve out(0 To outCount - 1)

                ' Position
                out(outCount - 8) = v(vi * 3)
                out(outCount - 7) = v(vi * 3 + 1)
                out(outCount - 6) = v(vi * 3 + 2)

                ' Normal (zero if not present)
                If ni >= 0 And vnCount > 0 Then
                    out(outCount - 5) = vn(ni * 3)
                    out(outCount - 4) = vn(ni * 3 + 1)
                    out(outCount - 3) = vn(ni * 3 + 2)
                Else
                    out(outCount - 5) = 0
                    out(outCount - 4) = 1
                    out(outCount - 3) = 0
                End If

                ' UV (zero if not present)
                If ti >= 0 And vtCount > 0 Then
                    out(outCount - 2) = vt(ti * 2)
                    out(outCount - 1) = vt(ti * 2 + 1)
                Else
                    out(outCount - 2) = 0
                    out(outCount - 1) = 0
                End If
            Next i
        End If
    Loop

    Close #f

    Dim result As OBJData
    result.vertices = out
    result.count = outCount / 8
    LoadOBJ = result
End Function

