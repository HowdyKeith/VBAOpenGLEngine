Option Explicit

' =========================================================
' CapsuleSweptCollider.cls (v8.1.1)
' =========================================================

Private m_EntityID As Long

Public Radius As Single
Public Height As Single
Public SkinWidth As Single
Public MaxSlopeAngle As Single
Public StepHeight As Single

' --- DAMPING SETTINGS ---
Public Friction As Single ' 0.9 = high friction, 0.99 = low friction
Public IsGrounded As Boolean
Public WasGrounded As Boolean

Private Sub Class_Initialize()
    m_EntityID = -1
    Radius = 0.3
    Height = 1.8
    SkinWidth = 0.05
    MaxSlopeAngle = 45#
    StepHeight = 0.4
    Friction = 0.9 ' Default 10% velocity loss per frame
    IsGrounded = False
    WasGrounded = False
End Sub

Public Sub Attach(ByVal entityID As Long): m_EntityID = entityID: End Sub

Public Sub Update(ByVal dt As Single, ByRef physics As Object)
    If m_EntityID < 0 Then Exit Sub
    WasGrounded = IsGrounded

    ' 1. Apply Damping (Friction) to X and Z only
    ComponentStore.VelocityX(m_EntityID) = ComponentStore.VelocityX(m_EntityID) * Friction
    ComponentStore.VelocityZ(m_EntityID) = ComponentStore.VelocityZ(m_EntityID) * Friction

    ' 2. Apply Gravity
    ComponentStore.VelocityY(m_EntityID) = ComponentStore.VelocityY(m_EntityID) + physics.Gravity * dt

    ' 3. Calculate Positions
    Dim px As Single, py As Single, pz As Single
    Dim prevX As Single, prevY As Single, prevZ As Single
    
    prevX = ComponentStore.PositionX(m_EntityID)
    prevY = ComponentStore.PositionY(m_EntityID)
    prevZ = ComponentStore.PositionZ(m_EntityID)

    px = prevX + ComponentStore.VelocityX(m_EntityID) * dt
    py = prevY + ComponentStore.VelocityY(m_EntityID) * dt
    pz = prevZ + ComponentStore.VelocityZ(m_EntityID) * dt

    ' 4. Physics Sweep
    Dim hit As Boolean, hitDist As Single, nx As Single, ny As Single, nz As Single
    hit = physics.SweepCapsule(prevX, prevY, prevZ, px - prevX, py - prevY, pz - prevZ, Radius, Height, hitDist, nx, ny, nz)

    If hit Then
        ' Simple collision resolve
        px = prevX + (px - prevX) * hitDist
        py = prevY + (py - prevY) * hitDist + SkinWidth
        pz = prevZ + (pz - prevZ) * hitDist

        ' Grounding check
        If ny > Cos(MaxSlopeAngle * 3.14159 / 180#) Then
            IsGrounded = True
            ComponentStore.VelocityY(m_EntityID) = 0
        End If
    Else
        IsGrounded = False
    End If

    ' 5. Write back
    ComponentStore.PositionX(m_EntityID) = px
    ComponentStore.PositionY(m_EntityID) = py
    ComponentStore.PositionZ(m_EntityID) = pz
End Sub

' --- Movement Logic ---
Public Sub Move(ByVal dx As Single, ByVal dz As Single)
    If m_EntityID < 0 Then Exit Sub
    ' We add force rather than setting it, so friction can pull it back down
    ComponentStore.VelocityX(m_EntityID) = ComponentStore.VelocityX(m_EntityID) + dx
    ComponentStore.VelocityZ(m_EntityID) = ComponentStore.VelocityZ(m_EntityID) + dz
End Sub


