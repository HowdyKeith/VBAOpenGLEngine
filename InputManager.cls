Option Explicit

' ============================================================
' InputManager.v7.1.cls
' Bridge between Win32 Events and Polling System
' ============================================================

Public InputSys As InputSystem

Private m_HasFocus As Boolean

Private Sub Class_Initialize()
    Set InputSys = New InputSystem
    m_HasFocus = True
End Sub

Public Sub Update()
    If Not InputSys Is Nothing Then InputSys.Update
End Sub

' ============================================================
' WIN32 EVENT PASSTHROUGHS
' ============================================================
Public Sub OnKeyDown(ByVal wParam As Long)
    If m_HasFocus Then InputSys.KeyDown wParam
End Sub

Public Sub OnKeyUp(ByVal wParam As Long)
    If m_HasFocus Then InputSys.KeyUp wParam
End Sub

Public Sub OnMouseMove(ByVal x As Long, ByVal y As Long)
    If m_HasFocus Then InputSys.MouseMove x, y
End Sub

' ============================================================
' PUBLIC ACCESSORS (USED BY GAME LOGIC)
' ============================================================
Public Function KeyDown(ByVal keyCode As Long) As Boolean
    KeyDown = InputSys.GetKey(keyCode)
End Function

Public Function KeyPressed(ByVal keyCode As Long) As Boolean
    KeyPressed = InputSys.GetKeyDown(keyCode)
End Function

Public Function KeyReleased(ByVal keyCode As Long) As Boolean
    KeyReleased = InputSys.GetKeyUp(keyCode)
End Function

Public Function MouseX() As Long: MouseX = InputSys.GetMouseX: End Function
Public Function MouseY() As Long: MouseY = InputSys.GetMouseY: End Function
Public Function MouseDeltaX() As Long: MouseDeltaX = InputSys.GetMouseDeltaX: End Function
Public Function MouseDeltaY() As Long: MouseDeltaY = InputSys.GetMouseDeltaY: End Function

' Handling Focus
Public Sub SetFocus(ByVal value As Boolean)
    m_HasFocus = value
End Sub


