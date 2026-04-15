Option Explicit

' =========================================================
' Module     : EmbeddedShaders
' Version    : v2.0  WEEK 3
' Description: Embedded GLSL shader library.
'              Week 3 adds: Star map, Spectra, Gas density shaders.
' =========================================================

Public Type ShaderEntry
    name      As String
    code      As String
    shaderType As Long
End Type

' ==================== BASIC PIPELINE ====================
Public Const BASIC_VERTEX As String = _
"#version 330 core" & vbCrLf & _
"layout (location = 0) in vec3 aPos;" & vbCrLf & _
"layout (location = 1) in vec3 aNormal;" & vbCrLf & _
"layout (location = 2) in vec2 aTexCoord;" & vbCrLf & _
"out vec3 FragPos; out vec3 Normal; out vec2 TexCoord;" & vbCrLf & _
"uniform mat4 model; uniform mat4 view; uniform mat4 projection;" & vbCrLf & _
"void main() {" & vbCrLf & _
"    FragPos = vec3(model * vec4(aPos, 1.0));" & vbCrLf & _
"    Normal = mat3(transpose(inverse(model))) * aNormal;" & vbCrLf & _
"    TexCoord = aTexCoord;" & vbCrLf & _
"    gl_Position = projection * view * vec4(FragPos, 1.0);" & vbCrLf & _
"}"

Public Const BASIC_FRAGMENT As String = _
"#version 330 core" & vbCrLf & _
"in vec3 FragPos; in vec3 Normal; in vec2 TexCoord;" & vbCrLf & _
"out vec4 FragColor;" & vbCrLf & _
"uniform sampler2D texture0;" & vbCrLf & _
"uniform vec3 lightDir = vec3(0.3, 0.7, 0.5); uniform vec3 viewPos;" & vbCrLf & _
"void main() {" & vbCrLf & _
"    vec3 norm = normalize(Normal);" & vbCrLf & _
"    float diff = max(dot(norm, normalize(lightDir)), 0.0);" & vbCrLf & _
"    vec3 viewDir = normalize(viewPos - FragPos);" & vbCrLf & _
"    vec3 reflectDir = reflect(-lightDir, norm);" & vbCrLf & _
"    float spec = pow(max(dot(viewDir, reflectDir), 0.0), 32.0);" & vbCrLf & _
"    vec3 ambient = vec3(0.25); vec3 diffuse = vec3(0.75)*diff; vec3 specular = vec3(0.6)*spec;" & vbCrLf & _
"    FragColor = vec4((ambient+diffuse+specular)*texture(texture0,TexCoord).rgb, 1.0);" & vbCrLf & _
"}"

' ==================== INSTANCED ====================
Public Const INSTANCED_VERTEX As String = _
"#version 330 core" & vbCrLf & _
"layout (location = 0) in vec3 aPos;" & vbCrLf & _
"layout (location = 1) in vec3 aNormal;" & vbCrLf & _
"layout (location = 2) in vec2 aTexCoord;" & vbCrLf & _
"layout (std430, binding = 0) buffer InstanceBlock { mat4 modelMatrices[]; };" & vbCrLf & _
"out vec3 FragPos; out vec3 Normal; out vec2 TexCoord;" & vbCrLf & _
"uniform mat4 view; uniform mat4 projection;" & vbCrLf & _
"void main() {" & vbCrLf & _
"    mat4 model = modelMatrices[gl_InstanceID];" & vbCrLf & _
"    FragPos = vec3(model * vec4(aPos, 1.0));" & vbCrLf & _
"    Normal = mat3(transpose(inverse(model))) * aNormal;" & vbCrLf & _
"    TexCoord = aTexCoord;" & vbCrLf & _
"    gl_Position = projection * view * model * vec4(aPos, 1.0);" & vbCrLf & _
"}"

' ==================== PARTICLES ====================
Public Const PARTICLE_VERTEX As String = _
"#version 330 core" & vbCrLf & _
"layout (location = 0) in vec3 aPos;" & vbCrLf & _
"uniform mat4 view; uniform mat4 projection;" & vbCrLf & _
"void main() {" & vbCrLf & _
"    gl_Position = projection * view * vec4(aPos, 1.0);" & vbCrLf & _
"    gl_PointSize = 5.0;" & vbCrLf & _
"}"

Public Const PARTICLE_FRAGMENT As String = _
"#version 330 core" & vbCrLf & _
"out vec4 FragColor;" & vbCrLf & _
"void main() { FragColor = vec4(1.0, 0.85, 0.4, 0.95); }"

' ==================== COMPUTE ====================
Public Const PARTICLE_COMPUTE As String = _
"#version 430 core" & vbCrLf & _
"layout (local_size_x = 256) in;" & vbCrLf & _
"struct Particle { vec4 position; vec4 velocity; float life; float padding[3]; };" & vbCrLf & _
"layout(std430, binding = 0) buffer ParticleBuffer { Particle particles[]; };" & vbCrLf & _
"uniform float deltaTime;" & vbCrLf & _
"void main() {" & vbCrLf & _
"    uint i = gl_GlobalInvocationID.x;" & vbCrLf & _
"    if (i >= particles.length()) return;" & vbCrLf & _
"    if (particles[i].life <= 0.0) {" & vbCrLf & _
"        particles[i].position = vec4(0,3,0,1);" & vbCrLf & _
"        particles[i].velocity = vec4((fract(sin(i*12.9898)*2.0-1.0)*12.0, 8.0+fract(sin(i))*18.0, (fract(sin(i*3.14))*2.0-1.0)*12.0, 0.0);" & vbCrLf & _
"        particles[i].life = 2.5+fract(sin(i*7.0))*3.5; return;" & vbCrLf & _
"    }" & vbCrLf & _
"    particles[i].position.xyz += particles[i].velocity.xyz * deltaTime;" & vbCrLf & _
"    particles[i].life -= deltaTime;" & vbCrLf & _
"}"

' ==================== WEEK 3: STAR MAP ====================
' Vertex: reads position (xyz), color (rgb), size from VBO.
' gl_PointSize driven by attribute so each star has correct apparent size.
' Requires GL_PROGRAM_POINT_SIZE enabled.
Public Const STAR_VERTEX As String = _
"#version 330 core" & vbCrLf & _
"layout (location = 0) in vec3  aPos;" & vbCrLf & _
"layout (location = 1) in vec3  aColor;" & vbCrLf & _
"layout (location = 2) in float aSize;" & vbCrLf & _
"out vec3  vColor;" & vbCrLf & _
"out float vAlpha;" & vbCrLf & _
"uniform mat4 view;" & vbCrLf & _
"uniform mat4 projection;" & vbCrLf & _
"uniform float screenH;" & vbCrLf & _
"void main() {" & vbCrLf & _
"    vec4 clip = projection * view * vec4(aPos, 1.0);" & vbCrLf & _
"    gl_Position = clip;" & vbCrLf & _
"    vColor = aColor;" & vbCrLf & _
"    // Size scales with projection (apparent magnitude effect)" & vbCrLf & _
"    float dist = length(clip.xyz);" & vbCrLf & _
"    gl_PointSize = clamp(aSize * screenH / max(dist, 0.1), 1.0, 32.0);" & vbCrLf & _
"    vAlpha = clamp(aSize * 2.0, 0.3, 1.0);" & vbCrLf & _
"}"

' Fragment: renders soft glow circle using gl_PointCoord.
' Discards corners so each point looks like a glowing star.
Public Const STAR_FRAGMENT As String = _
"#version 330 core" & vbCrLf & _
"in vec3  vColor;" & vbCrLf & _
"in float vAlpha;" & vbCrLf & _
"out vec4 FragColor;" & vbCrLf & _
"void main() {" & vbCrLf & _
"    vec2  coord = gl_PointCoord - vec2(0.5);" & vbCrLf & _
"    float r2    = dot(coord, coord);" & vbCrLf & _
"    if (r2 > 0.25) discard;" & vbCrLf & _
"    // Soft glow: bright core, fading halo" & vbCrLf & _
"    float core  = 1.0 - smoothstep(0.0,  0.04, r2);" & vbCrLf & _
"    float halo  = 1.0 - smoothstep(0.04, 0.25, r2);" & vbCrLf & _
"    float alpha = clamp(core * 1.0 + halo * 0.4, 0.0, 1.0) * vAlpha;" & vbCrLf & _
"    vec3  col   = mix(vColor, vec3(1.0), core * 0.6);" & vbCrLf & _
"    FragColor   = vec4(col, alpha);" & vbCrLf & _
"}"

' ==================== WEEK 3: SPECTRA ====================
' Simple coloured bar-chart shader for emission/absorption spectra.
' Each bar vertex carries its own pre-computed spectral colour.
Public Const SPECTRA_VERTEX As String = _
"#version 330 core" & vbCrLf & _
"layout (location = 0) in vec3 aPos;" & vbCrLf & _
"layout (location = 1) in vec3 aColor;" & vbCrLf & _
"out vec3 vColor;" & vbCrLf & _
"out float vHeight;" & vbCrLf & _
"uniform mat4 model;" & vbCrLf & _
"uniform mat4 view;" & vbCrLf & _
"uniform mat4 projection;" & vbCrLf & _
"void main() {" & vbCrLf & _
"    gl_Position = projection * view * model * vec4(aPos, 1.0);" & vbCrLf & _
"    vColor  = aColor;" & vbCrLf & _
"    vHeight = aPos.y;" & vbCrLf & _
"}"

Public Const SPECTRA_FRAGMENT As String = _
"#version 330 core" & vbCrLf & _
"in vec3  vColor;" & vbCrLf & _
"in float vHeight;" & vbCrLf & _
"out vec4 FragColor;" & vbCrLf & _
"uniform float maxHeight;" & vbCrLf & _
"uniform float glowStrength;" & vbCrLf & _
"void main() {" & vbCrLf & _
"    // Fade towards base, full brightness at peak" & vbCrLf & _
"    float t = clamp(vHeight / max(maxHeight, 0.001), 0.0, 1.0);" & vbCrLf & _
"    vec3 glow = vColor * (1.0 + glowStrength * t);" & vbCrLf & _
"    FragColor = vec4(glow, 0.9);" & vbCrLf & _
"}"

' ==================== WEEK 3: GAS DENSITY VOLUME ====================
' Billboard quad sprite for additive volume rendering.
' Each particle carries world position, density-derived colour and alpha.
Public Const VOLUME_VERTEX As String = _
"#version 330 core" & vbCrLf & _
"layout (location = 0) in vec3 aPos;" & vbCrLf & _
"layout (location = 1) in vec2 aUV;" & vbCrLf & _
"layout (location = 2) in vec4 aColor;" & vbCrLf & _
"out vec2 vUV;" & vbCrLf & _
"out vec4 vColor;" & vbCrLf & _
"uniform mat4 view;" & vbCrLf & _
"uniform mat4 projection;" & vbCrLf & _
"void main() {" & vbCrLf & _
"    gl_Position = projection * view * vec4(aPos, 1.0);" & vbCrLf & _
"    vUV    = aUV;" & vbCrLf & _
"    vColor = aColor;" & vbCrLf & _
"}"

Public Const VOLUME_FRAGMENT As String = _
"#version 330 core" & vbCrLf & _
"in vec2 vUV;" & vbCrLf & _
"in vec4 vColor;" & vbCrLf & _
"out vec4 FragColor;" & vbCrLf & _
"void main() {" & vbCrLf & _
"    // Gaussian soft-sphere blob" & vbCrLf & _
"    vec2  c = vUV - vec2(0.5);" & vbCrLf & _
"    float r = dot(c, c) * 4.0;" & vbCrLf & _
"    float a = vColor.a * exp(-r * 3.0);" & vbCrLf & _
"    FragColor = vec4(vColor.rgb * a, a);" & vbCrLf & _
"}"

' ============================================================
' EXPORT SYSTEM
' ============================================================
Public Sub ExportAllShaders(Optional ByVal folderName As String = "shaders", _
                            Optional ByVal forceOverwrite As Boolean = False)
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim p As String: p = ThisWorkbook.Path & "\" & folderName
    If Not fso.FolderExists(p) Then fso.CreateFolder p

    ExportShader fso, p & "\basic_vertex.glsl",     BASIC_VERTEX,     forceOverwrite
    ExportShader fso, p & "\basic_fragment.glsl",   BASIC_FRAGMENT,   forceOverwrite
    ExportShader fso, p & "\instanced_vertex.glsl", INSTANCED_VERTEX, forceOverwrite
    ExportShader fso, p & "\particle_vertex.glsl",  PARTICLE_VERTEX,  forceOverwrite
    ExportShader fso, p & "\particle_fragment.glsl",PARTICLE_FRAGMENT,forceOverwrite
    ExportShader fso, p & "\particle_compute.glsl", PARTICLE_COMPUTE, forceOverwrite
    ExportShader fso, p & "\star_vertex.glsl",      STAR_VERTEX,      forceOverwrite
    ExportShader fso, p & "\star_fragment.glsl",    STAR_FRAGMENT,    forceOverwrite
    ExportShader fso, p & "\spectra_vertex.glsl",   SPECTRA_VERTEX,   forceOverwrite
    ExportShader fso, p & "\spectra_fragment.glsl", SPECTRA_FRAGMENT, forceOverwrite
    ExportShader fso, p & "\volume_vertex.glsl",    VOLUME_VERTEX,    forceOverwrite
    ExportShader fso, p & "\volume_fragment.glsl",  VOLUME_FRAGMENT,  forceOverwrite

    Debug.Print "[EmbeddedShaders] Exported 12 shaders to: " & p
End Sub

Private Sub ExportShader(ByVal fso As Object, ByVal fullPath As String, _
                         ByVal content As String, ByVal forceOverwrite As Boolean)
    If forceOverwrite Or Not fso.FileExists(fullPath) Then
        Dim ts As Object
        Set ts = fso.CreateTextFile(fullPath, True)
        ts.Write content
        ts.Close
    End If
End Sub

Public Function GetAllShaders() As ShaderEntry()
End Function
