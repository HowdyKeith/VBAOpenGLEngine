Option Explicit
' ============================================================
' modMath.bas (FIXED v1.1)
' Matrix upload helpers using the GL facade and GLMath module.
' Original called bare glUniformMatrix4fv without the facade
' and undefined Identity/LookAt/Perspective functions.
' ============================================================

' ============================================================
' UPLOAD MODEL/VIEW/PROJECTION TO SHADER
' Uses GL.bas facade + GLMath.bas for matrix construction.
' ============================================================
Public Sub SetMatrices(ByVal program As Long, ByRef cam As Camera)
    ' Build matrices using GLMath module
    Dim model  As GLMath.Mat4
    Dim view   As GLMath.Mat4
    Dim proj   As GLMath.Mat4

    model = GLMath.Identity()

    view = GLMath.LookAt( _
        cam.Position.x, cam.Position.y, cam.Position.z, _
        cam.Position.x + 0, cam.Position.y + 0, cam.Position.z - 1, _
        0, 1, 0)

    proj = GLMath.Perspective(cam.FOV, cam.AspectRatio, cam.NearPlane, cam.FarPlane)

    ' Upload via GLUniform bridge (uses GL facade internally)
    Dim mModel As GLMatrix
    Dim mView  As GLMatrix
    Dim mProj  As GLMatrix
    Set mModel = New GLMatrix
    Set mView = New GLMatrix
    Set mProj = New GLMatrix

    Dim i As Long
    For i = 0 To 15
        mModel.m(i) = model.m(i)
        mView.m(i) = view.m(i)
        mProj.m(i) = proj.m(i)
    Next i

    GLUniform.SetMat4 program, "model", mModel
    GLUniform.SetMat4 program, "view", mView
    GLUniform.SetMat4 program, "projection", mProj
End Sub


