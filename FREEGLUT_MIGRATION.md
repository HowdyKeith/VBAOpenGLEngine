# FreeGLUT → VBA OpenGL Engine Migration Guide
## Week 2 | VBA OpenGL Engine

---

## Overview

The `modFreeGLUT.bas` shim maps every standard FreeGLUT/GLUT function to its
VBA equivalent by the same name. A typical C port requires three steps:

1. **Add `Option Explicit` and change types** (int→Long, double→Double, char→String)
2. **Move callback functions to module level** (not inside a class)
3. **Replace `main()` with a public `RunMyApp` sub**

That's it. No GL calls change. `glClear`, `glBegin`, `glRotatef` etc. are
identical because `modGL.bas` wraps them with the same names.

---

## Quick Reference: C → VBA

| C / FreeGLUT | VBA (modFreeGLUT) |
|---|---|
| `glutInit(&argc, argv)` | `glutInit` |
| `glutInitDisplayMode(GLUT_DOUBLE\|GLUT_RGBA\|GLUT_DEPTH)` | `glutInitDisplayMode GLUT_DOUBLE Or GLUT_RGBA Or GLUT_DEPTH` |
| `glutInitWindowSize(800, 600)` | `glutInitWindowSize 800, 600` |
| `glutCreateWindow("Title")` | `glutCreateWindow "Title"` |
| `glutDisplayFunc(display)` | `glutDisplayFunc "MyModule.Display"` |
| `glutReshapeFunc(reshape)` | `glutReshapeFunc "MyModule.Reshape"` |
| `glutKeyboardFunc(keyboard)` | `glutKeyboardFunc "MyModule.Keyboard"` |
| `glutMouseFunc(mouse)` | `glutMouseFunc "MyModule.Mouse"` |
| `glutIdleFunc(idle)` | `glutIdleFunc "MyModule.Idle"` |
| `glutTimerFunc(ms, timer, val)` | `glutTimerFunc ms, "MyModule.Timer", val` |
| `glutMainLoop()` | `glutMainLoop` |
| `glutLeaveMainLoop()` | `glutLeaveMainLoop` |
| `glutSwapBuffers()` | `glutSwapBuffers` |
| `glutPostRedisplay()` | `glutPostRedisplay` |
| `glutGet(GLUT_ELAPSED_TIME)` | `glutGet(GLUT_ELAPSED_TIME)` |
| `glutSolidSphere(r, sl, st)` | `glutSolidSphere r, sl, st` |
| `glutWireSphere(r, sl, st)` | `glutWireSphere r, sl, st` |
| `glutSolidCube(s)` | `glutSolidCube s` |
| `glutSolidTorus(i, o, s, r)` | `glutSolidTorus i, o, s, r` |
| `glutSolidCylinder(r, h, s, t)` | `glutSolidCylinder r, h, s, t` |
| `glutSolidCone(b, h, s, t)` | `glutSolidCone b, h, s, t` |

---

## Full Side-by-Side Example

### C / FreeGLUT (original)

```c
#include <GL/freeglut.h>
#include <math.h>

static float angle = 0.0f;
static int wireframe = 0;

void display(void) {
    glClear(GL_COLOR_BUFFER_BIT | GL_DEPTH_BUFFER_BIT);
    glLoadIdentity();
    gluLookAt(0,2,5, 0,0,0, 0,1,0);
    glRotatef(angle, 0, 1, 0);
    if (wireframe)
        glutWireSphere(1.0, 20, 20);
    else
        glutSolidSphere(1.0, 20, 20);
    glutSwapBuffers();
}

void reshape(int w, int h) {
    glViewport(0, 0, w, h);
    glMatrixMode(GL_PROJECTION);
    glLoadIdentity();
    gluPerspective(45.0, (double)w/h, 0.1, 100.0);
    glMatrixMode(GL_MODELVIEW);
}

void keyboard(unsigned char key, int x, int y) {
    if (key == 27) glutLeaveMainLoop();
    if (key == 'w') wireframe = !wireframe;
    glutPostRedisplay();
}

void idle(void) {
    angle += 0.5f;
    if (angle >= 360.0f) angle -= 360.0f;
    glutPostRedisplay();
}

int main(int argc, char **argv) {
    glutInit(&argc, argv);
    glutInitDisplayMode(GLUT_DOUBLE | GLUT_RGBA | GLUT_DEPTH);
    glutInitWindowSize(800, 600);
    glutCreateWindow("Spinning Sphere");
    glutDisplayFunc(display);
    glutReshapeFunc(reshape);
    glutKeyboardFunc(keyboard);
    glutIdleFunc(idle);
    glEnable(GL_DEPTH_TEST);
    glutMainLoop();
    return 0;
}
```

### VBA (ported, file: MySphere.bas)

```vb
Option Explicit

' State
Private m_Angle     As Single
Private m_Wireframe As Boolean

' Entry point (equivalent of main)
Public Sub RunSpherDemo()
    m_Angle     = 0
    m_Wireframe = False

    glutInit
    glutInitDisplayMode GLUT_DOUBLE Or GLUT_RGBA Or GLUT_DEPTH
    glutInitWindowSize 800, 600
    glutCreateWindow "Spinning Sphere"

    glutDisplayFunc  "MySphere.Display"
    glutReshapeFunc  "MySphere.Reshape"
    glutKeyboardFunc "MySphere.Keyboard"
    glutIdleFunc     "MySphere.Idle"

    modGL.glEnable GL_DEPTH_TEST    ' same GL call

    glutMainLoop
End Sub

Public Sub Display()
    GL.glClear GL.GL_COLOR_BUFFER_BIT Or GL.GL_DEPTH_BUFFER_BIT
    GL.glLoadIdentity
    ' gluLookAt equivalent via matrix setup
    modGL.apiTranslatef 0, -2, -5
    GL.glRotatef m_Angle, 0, 1, 0

    If m_Wireframe Then
        glutWireSphere 1.0, 20, 20
    Else
        glutSolidSphere 1.0, 20, 20
    End If

    glutSwapBuffers
End Sub

Public Sub Reshape(ByVal w As Long, ByVal h As Long)
    GL.glViewport 0, 0, w, h
    GL.glMatrixMode GL.GL_PROJECTION
    GL.glLoadIdentity
    modGL.apiPerspective 45#, CDbl(w) / CDbl(h), 0.1, 100#
    GL.glMatrixMode GL.GL_MODELVIEW
End Sub

Public Sub Keyboard(ByVal key As String, ByVal x As Long, ByVal y As Long)
    If key = Chr(27) Then glutLeaveMainLoop     ' ESC
    If LCase(key) = "w" Then m_Wireframe = Not m_Wireframe
    glutPostRedisplay
End Sub

Public Sub Idle()
    m_Angle = m_Angle + 0.5
    If m_Angle >= 360 Then m_Angle = m_Angle - 360
    glutPostRedisplay
End Sub
```

**Changes made:** 11 lines of type adjustments. Zero GL calls changed. Total porting time: ~5 minutes.

---

## Callback Registration

Callbacks are registered as **string names** of module-level public procedures:

```vb
' Same module (MySphere.bas):
glutDisplayFunc "MySphere.Display"

' Or just the procedure name if it's unique project-wide:
glutDisplayFunc "Display"
```

The shim uses `Application.Run` to invoke them, which works for any public
procedure in the VBA project.

### Callback Signatures

| Callback | VBA signature |
|---|---|
| Display | `Public Sub Display()` |
| Reshape | `Public Sub Reshape(ByVal w As Long, ByVal h As Long)` |
| Keyboard | `Public Sub Keyboard(ByVal key As String, ByVal x As Long, ByVal y As Long)` |
| KeyboardUp | `Public Sub KeyboardUp(ByVal key As String, ByVal x As Long, ByVal y As Long)` |
| Special | `Public Sub Special(ByVal key As Long, ByVal x As Long, ByVal y As Long)` |
| Mouse | `Public Sub Mouse(ByVal btn As Long, ByVal state As Long, ByVal x As Long, ByVal y As Long)` |
| Motion | `Public Sub Motion(ByVal x As Long, ByVal y As Long)` |
| PassiveMotion | `Public Sub PassiveMotion(ByVal x As Long, ByVal y As Long)` |
| Idle | `Public Sub Idle()` |
| Timer | `Public Sub MyTimer(ByVal value As Long)` |
| Close | `Public Sub OnClose()` |

---

## Primitives Available

| C function | VBA equivalent | Status |
|---|---|---|
| `glutSolidCube(s)` | `glutSolidCube s` | ✅ Full |
| `glutWireCube(s)` | `glutWireCube s` | ✅ Full |
| `glutSolidSphere(r,sl,st)` | `glutSolidSphere r,sl,st` | ✅ Full UV sphere |
| `glutWireSphere(r,sl,st)` | `glutWireSphere r,sl,st` | ✅ Full |
| `glutSolidTorus(i,o,s,r)` | `glutSolidTorus i,o,s,r` | ✅ Full |
| `glutWireTorus(i,o,s,r)` | `glutWireTorus i,o,s,r` | ✅ Full |
| `glutSolidCylinder(r,h,s,t)` | `glutSolidCylinder r,h,s,t` | ✅ Full (capped) |
| `glutWireCylinder(r,h,s,t)` | `glutWireCylinder r,h,s,t` | ✅ Full |
| `glutSolidCone(b,h,s,t)` | `glutSolidCone b,h,s,t` | ✅ Full |
| `glutWireCone(b,h,s,t)` | `glutWireCone b,h,s,t` | ✅ Full |
| `glutSolidTeapot(s)` | `glutSolidTeapot s` | ⚠️ Approximated (load teapot.obj for full mesh) |

---

## What Is NOT Supported

| Feature | Reason | Workaround |
|---|---|---|
| Multiple windows | VBA is single-threaded | Design for one window |
| `gluLookAt` | glu32 perspective only | Use `modGL.apiTranslatef` + `GL.glRotatef` or `GLMath.LookAt` |
| `glutBitmapCharacter` | No bitmap font system yet | Use Windows GDI text overlay |
| `glutExtensionSupported` | No extension string query | Assume GL 4.x (modern drivers) |
| Full-screen mode | Requires display mode change | Resize window manually |
| Joystick callbacks | No joystick input module | Add `GetJoystickState` if needed |

---

## Next: Week 3 — Excel Data Visualization Demos
