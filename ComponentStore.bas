Option Explicit

' ==============================================================================
' Module     : ComponentStore
' Version    : v8.11
' Description: ECS Global Data Store with Entity State management.
' ==============================================================================

Public entityCount As Long

' --- ENTITY STATE ---
Public IsActive() As Boolean
Public IsStatic() As Boolean

' --- TRANSFORM (PRS) ---
Public PositionX() As Single, PositionY() As Single, PositionZ() As Single
Public RotationX() As Single, RotationY() As Single, RotationZ() As Single
Public ScaleX() As Single, ScaleY() As Single, ScaleZ() As Single

' --- PHYSICS (Velocity) ---
Public VelocityX() As Single, VelocityY() As Single, VelocityZ() As Single

' --- SPATIAL (AABB) ---
Public AabbMinX() As Single, AabbMinY() As Single, AabbMinZ() As Single
Public AabbMaxX() As Single, AabbMaxY() As Single, AabbMaxZ() As Single

' --- RENDERING ---
Public MeshVAO() As Long        ' The Container
Public MeshVBO() As Long        ' The Raw Vertex Buffer <--- ADDED
Public MeshIBO() As Long        ' The Index Buffer <--- ADDED
Public MeshIndexCount() As Long
Public MaterialID() As Long

Public Sub ECS_InitStore()
    entityCount = 0
    Erase IsActive: Erase IsStatic
    Erase PositionX: Erase PositionY: Erase PositionZ
    Erase RotationX: Erase RotationY: Erase RotationZ
    Erase ScaleX: Erase ScaleY: Erase ScaleZ
    Erase VelocityX: Erase VelocityY: Erase VelocityZ
    Erase AabbMinX: Erase AabbMinY: Erase AabbMinZ
    Erase AabbMaxX: Erase AabbMaxY: Erase AabbMaxZ
    Erase MeshVAO: Erase MeshVBO: Erase MeshIBO: Erase MeshIndexCount: Erase MaterialID
End Sub

Public Function ECS_CreateEntity() As Long
    Dim i As Long
    i = entityCount: entityCount = entityCount + 1
    
    ReDim Preserve IsActive(i): ReDim Preserve IsStatic(i)
    ReDim Preserve PositionX(i): ReDim Preserve PositionY(i): ReDim Preserve PositionZ(i)
    ReDim Preserve RotationX(i): ReDim Preserve RotationY(i): ReDim Preserve RotationZ(i)
    ReDim Preserve ScaleX(i): ReDim Preserve ScaleY(i): ReDim Preserve ScaleZ(i)
    ReDim Preserve VelocityX(i): ReDim Preserve VelocityY(i): ReDim Preserve VelocityZ(i)
    ReDim Preserve AabbMinX(i): ReDim Preserve AabbMinY(i): ReDim Preserve AabbMinZ(i)
    ReDim Preserve AabbMaxX(i): ReDim Preserve AabbMaxY(i): ReDim Preserve AabbMaxZ(i)
    
    ReDim Preserve MeshVAO(i)
    ReDim Preserve MeshVBO(i)        ' <--- ADDED
    ReDim Preserve MeshIBO(i)        ' <--- ADDED
    ReDim Preserve MeshIndexCount(i)
    ReDim Preserve MaterialID(i)
    
    ' Defaults
    IsActive(i) = True
    IsStatic(i) = False
    ScaleX(i) = 1!: ScaleY(i) = 1!: ScaleZ(i) = 1!
    MaterialID(i) = -1
    ECS_CreateEntity = i
End Function


