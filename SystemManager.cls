' SystemManager.cls (v1.0 ECS HARDENED)
Option Explicit

Private m_Systems As Object
Private m_Order As Collection
Private m_Enabled As Object

Private m_World As Object

Private m_LastTime As Double
Private m_DeltaTime As Single

Private Sub Class_Initialize()

    Set m_Systems = CreateObject("Scripting.Dictionary")
    Set m_Order = New Collection
    Set m_Enabled = CreateObject("Scripting.Dictionary")

    m_LastTime = Timer

End Sub

' =========================================================
' WORLD LINK
' =========================================================

Public Sub ECS_Initialize(ByVal World As Object)
    Set m_World = World
End Sub

' =========================================================
' REGISTER SYSTEM
' =========================================================

Public Sub ECS_RegisterSystem(ByVal name As String, ByVal sys As Object)

    If Not m_Systems.Exists(name) Then
        m_Systems.Add name, sys
        m_Order.Add name
        m_Enabled.Add name, True
    End If

End Sub

' =========================================================
' MAIN UPDATE LOOP
' =========================================================

Public Sub SYS_Tick()

    Dim currentTime As Double
    currentTime = Timer

    m_DeltaTime = currentTime - m_LastTime
    m_LastTime = currentTime

    If m_DeltaTime > 0.25 Then m_DeltaTime = 0.25

    Dim n As Variant
    Dim sys As Object

    For Each n In m_Order

        If m_Enabled(n) Then

            Set sys = m_Systems(n)

            CallByName sys, "SYS_Update", VbMethod, m_World, m_DeltaTime

            CallByName sys, "SYS_LateUpdate", VbMethod, m_World, m_DeltaTime

        End If

    Next n

End Sub

' =========================================================
' RENDER PHASE
' =========================================================

Public Sub SYS_Render()

    Dim n As Variant
    Dim sys As Object

    For Each n In m_Order

        If m_Enabled(n) Then

            Set sys = m_Systems(n)

            CallByName sys, "SYS_Render", VbMethod, m_World

        End If

    Next n

End Sub

' =========================================================
' DELTA TIME
' =========================================================

Public Function SYS_DeltaTime() As Single
    SYS_DeltaTime = m_DeltaTime
End Function


