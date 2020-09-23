VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cDX8"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'RVTDX.cDX8.cls - a component of the rvtDX.dll
'©2003 Ron van Tilburg - rivit@f1.net.au
'Freeware for Educational Purposes, For commercial interests contact author please, I retain copyright.

'cDX8.cls - The Root Class,  All setup and destruction goes through here

'Supported Standard Figures - defined Here because its needed in a number of places

Public Enum FIG_SPEC    'Defined in World Space By: All defined shapes Fit a unit cube, centred at 0,0,0

  'FigSpec is made up as a long P3,P2,P1,TT
  'This is handled automatically in the cFigure.SetFigSpec Method
  FIGS_TYPEMASK = &H7F&     'TYPE 1-127
  FIGS_P1MASK = &HFF00&     'Parameter 1  'If parameter is a fraction Parm/256 is used
  FIGS_P2MASK = &HFF0000    'Paremater 2
  FIGS_P3MASK = &HFF000000  'Parameter 3

  'The Camera, ViewPoint and Other Markers are Special Figures without a VertexSet
  FIGS_LOCATOR = 0          'xyz - a point without rendering                                  0 Vertex

  'The Axes, and AxisPlanes are Special Figures
  FIGS_AXES = &H1           'xyz - three lines (compounded with labels)                       40 Vertices
  FIGS_AXISPLANES = &H101&  'xyz -                                                           1492 vertices

  'MESH Figures
  FIGS_MESH = 2             'Generated Mesh
  FIGS_TEXT = &H102&        'Using Figure Text and TTF Font                                    N vertices
  FIGS_TEAPOT = &H202&      'Uses DX8 Teapot                                                   N vertices
  
  'Planar Lineal Figures    'xyz
  FIGS_POINT = 3            'a point                                                           1 Vertex
  FIGS_LINE = 4             'two points                                                        2 vertices
  FIGS_POLYLINE = 5         'P1=N N>3                                                          n vertices
  FIGS_POLYARC = 6          'P1=N vertices, P2= Angle (0-255)/256*360                          N vertices

  'Planar Areal Figures     'x0z
  FIGS_POLYGON = 7          'P1=N vertices n>=3                                                N vertices
  FIGS_POLYWASHER = 8       'P1=N vertices n>=3, P2=Inner Radius  kk/256                      2N vertices
  FIGS_POLYSTAR = 9         'P1=N vertices n>=3, P2=Pointiness 0 = Normal                     2N vertices
  FIGS_POLYWEDGE = 10       'P1=N vertices, P2 = included angle kk/256*360
  FIGS_POLYMESH = 11        'P1=N vertex shape, n=3,4,5,6,8, P2=M the number in any given direction

  'Solid Figures            'xyz
  FIGS_REGULARSOLID = 12    'P1=4,6,8,12,20 ONLY
  FIGS_SPHEROID = 13        'P1=N n>=3 Longitude Slices,P2=M m>=3 Latitude slices               N vertices
  FIGS_PRISM = 14           'P1=N n>=3, 2 equal ends joined in space                            N vertices
  FIGS_FRUSTRUM = 15        'P1=N n>=3, P2=Radius Top, 2 ends joined in space                   N vertices
  FIGS_TOROID = 16          'P1=N n>=3, P2=M n>3, P3=Inner Radius                               N vertices
  FIGS_PIPE = 17            'P1=N n>=3, P2=Inner Radius                                         N vertices
  FIGS_ASTROID = 18         'P1=N n>=3  P2=M n>=3 starlike - pointy bits                        N vertices
  FIGS_NEBULA = 19          'P1=N n>=3                                                          N vertices
  FIGS_SNAIL = 20           'P1=N n>=3                                                          N vertices
  FIGS_BEAD = 21            'P1=N n>=3                                                          N vertices
  FIGS_POPE = 22            'P1=N n>=3                                                          N vertices
  FIGS_WALNUT = 23          'P1=N n>=3                                                          N vertices
  FIGS_PERTSPHERE = 24      'P1=N n>=3                                                          N vertices
  FIGS_FLOWER = 25          'P1=N n>=3                                                          N vertices
  FIGS_SHELL = 26           'P1=N n>=3                                                          N vertices
  FIGS_LILY = 27            'P1=N n>=3                                                          N vertices

  FIGS_POINTFIELD = 28      'a pointfield                                                       N Vertices
  FIGS_POINTSPHERE = 29     'a spherical pointfield                                             N Vertices
  FIGS_SHEET = 30           'a simple plane, P1=n waves P2=m waves,P3=Ratio Random     (N+1)(M+1) Vertices
  
  'FILL NEW ONES IN HERE -v--- and Modify FIG_LAST, and The Generator ---
  '---------------------------------------------------
  '---------------------------------------------------
  FIGS_FIRST = 1
  FIGS_LAST = 30            'Highest value allowed is 127  &H7F

  'User defined, custome generated Figures should go through Here  &H81-&HFF
  FIGS_CUSTOM = &H80&       'IF SET a User Class Generates the Vertex and Index Buffer, else use Stock Object

  'Special Definitions already worked out
  'Named Lineal Figures
  FIGS_CIRCLE = &H2405&      'PolyLine n=36

  'Named Areal Figures
  FIGS_TRIANGLE = &H307&     'Poly n=3 = Equilateral                                            3  vertices
  FIGS_SQUARE = FIGS_SHEET   'Sheet n=0 = Square                                                4  vertices
  FIGS_PENTAGON = &H507&     'Poly n=5 = Pentagon                                               4  vertices
  FIGS_HEXAGON = &H607&      'Poly n=6 = Hexagon                                                4  vertices
  FIGS_OCTAGON = &H807&      'Poly n=8 = Octagon                                                4  vertices
  FIGS_DISK = &H2407&        'a flat filled circle   P1=36                                      N  vertices
  FIGS_WASHER = &H802408     'Washer P1=36 n>=3, P2 = 128/256=0.5                               N  vertices
  FIGS_STAR = &H509&         'A standard 5 Pointed Star  P1=5,P2=0                             30  vertices
  FIGS_SEMIDISK = &H80120A   'Wedge P1=18,P2=128/256*360deg                                   1+n  vertices
  FIGS_PLANE = &H40B&        'Poly P1=4, P2=N, P3=M = Plane of NxM  squares                   4nm  Vertices

  'Named Regular Solids
  FIGS_TETRAHEDRON = &H40C&     'P1=4 four equal triangles                                       12 vertices
  FIGS_CUBE = &H60C&            'P1=6 squares                                                    36 vertices
  FIGS_OCTAHEDRON = &H80C&      'P1=8 8 equilateral triangles                                    24 vertices
  FIGS_DODECAHEDRON = &HC0C&    'P1=12 12 pentagons                                              32 vertices
  FIGS_DODECAHEDRON2 = &HD0C&   'P1=13 12 pentagons, but recessed centres                        32 vertices
  FIGS_ICOSAHEDRON = &H140C&    'P1=20 20 eqilateral triangles                                   24 vertices
  FIGS_HEXADECAHEDRON = &H3C0C& 'P1=60 12 pentagons as 60 isosceles triangles                    32 vertices

  'Named Solids
  FIGS_SPHERE = &H12240D     'Spheroid P1=36,P2=18                                               N vertices
  FIGS_CYLINDER = &H240E&    'Prism P1=36                                                        N vertices
  FIGS_CONE = &H240F&        'Frustrum P1=36,P2=0                                                N vertices
  FIGS_HALFCONE = &H80240F   'Prism P1=36,P2=0.5                                                 N vertices
  FIGS_QTRCONE = &HC0240F    'Prism P1=36,P2=0.75                                                N vertices
  FIGS_TORUS = &H80242410    'Toroid P1=36, P2=36, P3=0.5                                        N vertices
  FIGS_ASTER = &H122412      'Astroid P1=36,P2=18                                                N vertices
End Enum

Public Enum RENDER_FLAGS
  RF_DRAWSOLID = &H0              'Show Solid, Lit, With Materials and Textures, Gouraud Shading, All Vertices
  RF_DRAWPOINTS = &H1             'Show in Point Mode
  RF_DRAWWIREFRAME = &H2          'Show in WireFrame Mode
  
  RF_UNLIT = &H4                  'No Lighting At all
  RF_VERTEXCOLOURS = &H8 Or &H4   'Must be unlit to use Vertex Colours on their own
  RF_NOTEXTURES = &H10            'No Textures At all
  RF_SHINY = &H20                 'Render In Specular Mode otherwise dull
  
  RF_LIGHTTINT = &H40
  RF_DARKTINT = &H80
  RF_TRANSPARENT = &H100
  
  'RF_INVERT = &H200               'cant figure this out
  
  RF_SHADEFLAT = &H400            'Flat Shading otherwise Gouraud
  RF_REMOVEHIDDEN = &H800         'CULL hidden Surfaces otherwise not
  
  RF_IGNORE = &H80000000          'Dont Render THIS one at all
  
  RF_ALPHAOPS = RF_LIGHTTINT Or RF_DARKTINT Or RF_TRANSPARENT     'Only used internally
End Enum

'The Vertices used Throughout
Public Type ZVertex      'CUSTOM
  x  As Single    'Coord
  y  As Single
  z  As Single
  nx As Single    'normals
  ny As Single
  nz As Single
  dc As Long      'diffuse color
  ut As Single    'textures
  vt As Single
End Type

'=========================================================================================================
'============================== DX8 SETUP ================================================================
'=========================================================================================================

Public Function InitDX8(ByVal VBhWnd As Long) As Boolean    'Bind to This Level 1 Window
Attribute InitDX8.VB_Description = "Prepare for using DirectX"

' Set the global DirectX objects of DirectInput, DirectSound, and Direct3D
' Return True if we survive

  Set DX8 = New DirectX8       ' Get an Instance of DirectX8

  If DX8 Is Nothing Then
    ErrorMsgBox "DirectX8 appears to be missing - I can't continue", Err, vbCritical
    Exit Function
  End If  'Whew!! we seem to have DX8

  'Get DirectInput Setup
  Set DI8 = DX8.DirectInputCreate()
  If Not DI8 Is Nothing Then
    Set Keyboard = DI8.CreateDevice("GUID_SysKeyboard")   'Get Keyboard
    If Keyboard Is Nothing Then
      ErrorMsgBox "DirectInput.Keyboard cannot be loaded - I can't continue", Err, vbCritical
      Exit Function
    End If

    Set Mouse = DI8.CreateDevice("guid_SysMouse")         'Get Mouse
    If Mouse Is Nothing Then
      ErrorMsgBox "DirectInput.Mouse cannot be loaded - I can't continue", Err, vbCritical
      Exit Function
    End If

    ' Set the data formats to the commmonly used keyboard and mouse
    Call Keyboard.SetCommonDataFormat(DIFORMAT_KEYBOARD)
    Call Mouse.SetCommonDataFormat(DIFORMAT_MOUSE)

    ' Set cooperative level, this tells DI how much control we need
    Call Keyboard.SetCooperativeLevel(VBhWnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND)
    Call Mouse.SetCooperativeLevel(VBhWnd, DISCL_NONEXCLUSIVE Or DISCL_BACKGROUND)

    ' Make available (bind) to our program
    Call Keyboard.Acquire
    Call Mouse.Acquire

  Else
    ErrorMsgBox "DirectInput cannot be loaded - I can't continue", Err, vbCritical
    Exit Function
  End If

  'Get DirectSound Setup
  Set DS8 = DX8.DirectSoundCreate("")                   'Getting the Prime Sound Object
  If Not DS8 Is Nothing Then
    Call DS8.SetCooperativeLevel(VBhWnd, DSSCL_NORMAL)    'Just basic control
  Else
    ErrorMsgBox "DirectSound cannot be loaded - I can't continue", Err, vbCritical
    Exit Function
  End If

  'Get Direct3D Setup
  Set D3D8 = DX8.Direct3DCreate()                        'Getting the Prime Graphics Object
  If D3D8 Is Nothing Then
    ErrorMsgBox "Direct3D cannot be loaded - I can't continue", Err, vbCritical
    Exit Function
  End If

  Set D3DX = New D3DX8                                  'Getting the Helper Graphics Object
  If D3DX Is Nothing Then
    ErrorMsgBox "Direct3DX cannot be loaded - I can't continue", Err, vbCritical
    Exit Function
  End If

  'The constant matrices
  WORLD_YISUP.y = 1

  InitDX8 = True        'we made it
  DXhWnd = VBhWnd

End Function

'---------------------------------------- SHUTDOWN DX8 --------------------------------------------------

Public Sub ShutDownDX8()    'unloads all the DirectX objects - THIS MUST BE CALLED BEFORE CLOSING PROGRAM
Attribute ShutDownDX8.VB_Description = "Kill off DirectX"

  On Error Resume Next  'Deletion In Reverse Order

  Set D3DD = Nothing
  Set D3DX = Nothing
  Set D3D8 = Nothing

  Call PlaySoundFromFile("")    'does  Set Sound = Nothing
  Set DS8 = Nothing

  Set Keyboard = Nothing
  Set Mouse = Nothing
  Set DI8 = Nothing
  Set DX8 = Nothing
  DXhWnd = 0                'Drop This Link

  On Error GoTo 0

End Sub

'-------------------------------------- PRE RENDER STEPS ---------------------------------------------

'Call This after InitDX8, and Before Rendering

'InVBWindow - if True just use the VB Window
'Width  - IF Not InVBWIndow
'Height - IF Not InVBWIndow
'Depth  - IF Not InVBWIndow (8),16,32   'Bit Depth

Public Function GetRenderDevice(Optional ByVal InVBWindow As Boolean = False, _
                                Optional ByVal Width As Long = 640, _
                                Optional ByVal Height As Long = 480, _
                                Optional ByVal Depth As Long = 16) As Boolean
Attribute GetRenderDevice.VB_Description = "Build a DirectX rendering Device"

 Dim D3DP         As D3DPRESENT_PARAMETERS       'How DirectX will Present (display) it's rendering
 Dim DevType      As CONST_D3DDEVTYPE
 Dim DevVxProc    As Long

  'Fill the D3DP (D3DPRESENT_PARAMETERS), describing how DirectX should display
  If InVBWindow Then
    D3DP.Windowed = 1
    Call D3D8.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, D3DDM)    ' get current display mode
    Call GetWindowRect(DXhWnd, DXHWndRect)                        ' get the area we will work with
    D3DDM.Width = DXHWndRect.Right - DXHWndRect.Left
    D3DDM.Height = DXHWndRect.bottom - DXHWndRect.Top
  Else
    D3DP.Windowed = 0
    D3DDM.Format = CheckDisplayMode(Width, Height, Depth)        ' check if requested mode is supported
    If D3DDM.Format = D3DFMT_UNKNOWN Then                        ' it isnt so use current
      Call D3D8.GetAdapterDisplayMode(D3DADAPTER_DEFAULT, D3DDM)  ' get current display mode
      MsgBox "Display mode: " _
             & Width & "x" & Height & "x" & Depth & "isnt supported" & vbCrLf _
             & "I will carry on with Display Mode :" _
             & D3DDM.Width & "x" & D3DDM.Height & "x" & D3DDM.Format, vbInformation
    Else
      D3DDM.Width = Width
      D3DDM.Height = Height
    End If
  End If

  'Common Parameters
  With D3DP
    .BackBufferFormat = D3DDM.Format
    .BackBufferWidth = D3DDM.Width
    .BackBufferHeight = D3DDM.Height
    .EnableAutoDepthStencil = 1
    .AutoDepthStencilFormat = D3DFMT_D16            'Use z Depth sorting
    .BackBufferCount = 1                            'just render and display
    .SwapEffect = D3DSWAPEFFECT_COPY_VSYNC
  End With

  'Check what device capabilities exist
  On Local Error Resume Next
  DevType = D3DDEVTYPE_HAL
  Call D3D8.GetDeviceCaps(D3DADAPTER_DEFAULT, DevType, D3DDC)
  If Err.Number = D3DERR_INVALIDDEVICE Then                       'We couldn't use the hardware device
    DevType = D3DDEVTYPE_REF
    Call D3D8.GetDeviceCaps(D3DADAPTER_DEFAULT, DevType, D3DDC)   'use the reference device
    Err.Clear
  End If

  'Check vertex processing capabilities
  If (D3DDC.DevCaps And D3DDEVCAPS_HWTRANSFORMANDLIGHT) <> 0 Then         'Hardware can do it
    DevVxProc = D3DCREATE_HARDWARE_VERTEXPROCESSING                       'so we will use it
  Else
    DevVxProc = D3DCREATE_SOFTWARE_VERTEXPROCESSING                       'the CPU will do it
  End If

  '  Call ListDeviceCaps(DevCaps)     'UNCOMMENT THIS TO PRINT YOUR ADAPTER'S CAPABILITIES TO FILE

  '//For Pure Device processing:
  'If (DevCaps.DevCaps And D3DDEVCAPS_PUREDEVICE) Then
  '  DevVxProc = D3DCREATE_PUREDEVICE
  'Else
  '    Debug.Print "Pure device is not supported"
  'End If

  On Local Error Resume Next
  Set D3DD = D3D8.CreateDevice(D3DADAPTER_DEFAULT, DevType, DXhWnd, DevVxProc, D3DP)
  If D3DD Is Nothing Then
    ErrorMsgBox "Direct3D cannot make a render device - I can't continue", Err, vbCritical
    Exit Function
  End If

  ' Set rendering defaults
  With D3DD
    Call .SetRenderState(D3DRS_CULLMODE, D3DCULL_NONE)      ' Dont CULL Hidden Planes
    Call .SetRenderState(D3DRS_ZENABLE, D3DZB_TRUE)         ' enable z buffering
    Call .SetRenderState(D3DRS_ZBIAS, 2)                    ' Keep coplanar things right side of each other
    Call .SetRenderState(D3DRS_NORMALIZENORMALS, 0)         ' dont renormalise normals (faster, but better??)
    Call .SetRenderState(D3DRS_FILLMODE, D3DFILL_SOLID)                    ' render solid polygons
    Call .SetRenderState(D3DRS_LIGHTING, 1)                                ' enable lighting
    Call .SetRenderState(D3DRS_COLORVERTEX, 1)                             ' enable vertex colours
    Call .SetRenderState(D3DRS_AMBIENT, &HFFFFFF)                          ' use ambient Grey light
    Call .SetRenderState(D3DRS_SHADEMODE, D3DSHADE_GOURAUD)                ' Use Gouraud Shading
    Call .SetRenderState(D3DRS_SPECULARMATERIALSOURCE, D3DMCS_MATERIAL)    'use material specular
  End With

  GetRenderDevice = True    'it worked !
  On Error GoTo 0

Exit Function

ErrHandler:
  Call GeneralErrorMsg(Err)
  On Error GoTo 0

End Function

Public Sub ListDeviceCaps(ByRef DevCaps As D3DCAPS8)  'What the D3DDevice is capable of
Attribute ListDeviceCaps.VB_Description = "List to File the capabilities of your Graphics Card"

  Open App.Path & "\D3DDCaps.txt" For Output As 1
  With DevCaps
    Print #1, "D3D Device Capabilities"
    Print #1, "Adapter Ordinal = "; .AdapterOrdinal
    Print #1, "AlphaCmpCaps    = "; Hex8(.AlphaCmpCaps)
    Print #1, "Caps            = "; Hex8(.Caps)
    Print #1, "Caps2           = "; Hex8(.Caps2)
    Print #1, "Caps3           = "; Hex8(.Caps3)
    Print #1, "CubeTexFltCaps  = "; Hex8(.CubeTextureFilterCaps)
    Print #1, "Cursor Caps     = "; Hex8(.CursorCaps)
    Print #1, "DestBlendCaps   = "; Hex8(.DestBlendCaps)
    Print #1, "DevCaps         = "; Hex8(.DevCaps)
    Print #1, "DevType         = "; Hex8(.DeviceType)
    Print #1, "Extent Adjust   = "; Hex8(.ExtentsAdjust)
    Print #1, "FVFCaps         = "; Hex8(.FVFCaps)
    Print #1, "Guard Band T    = "; .GuardBandTop
    Print #1, "Guard Band L    = "; .GuardBandLeft
    Print #1, "Guard Band B    = "; .GuardBandBottom
    Print #1, "Guard Band R    = "; .GuardBandRight
    Print #1, "LineCaps        = "; Hex8(.LineCaps)
    Print #1, "Max Act Lights  = "; .MaxActiveLights
    Print #1, "Max Anisotropy  = "; .MaxAnisotropy
    Print #1, "Max Pixel Shade = "; .MaxPixelShaderValue
    Print #1, "Max Point Size  = "; .MaxPointSize
    Print #1, "Max Primitives  = "; .MaxPrimitiveCount
    Print #1, "Max Sim Textures= "; .MaxSimultaneousTextures
    Print #1, "Max Streams     = "; .MaxStreams
    Print #1, "Max Stream Strid= "; .MaxStreamStride
    Print #1, "Max Tex Asp Rat.= "; .MaxTextureAspectRatio
    Print #1, "Max Tex Blends  = "; .MaxTextureBlendStages
    Print #1, "Max Tex Height  = "; .MaxTextureHeight
    Print #1, "Max Tex Repeat  = "; .MaxTextureRepeat
    Print #1, "Max Tex Width   = "; .MaxTextureWidth
    Print #1, "Max Usr Clip Pl = "; .MaxUserClipPlanes
    Print #1, "Max Vx Blnd Mats= "; .MaxVertexBlendMatrices
    Print #1, "Max Vx BLnd MIdx= "; .MaxVertexBlendMatrixIndex
    Print #1, "Max Vx Index    = "; .MaxVertexIndex
    Print #1, "Max Vx Shade Cnt= "; .MaxVertexShaderConst
    Print #1, "Max Vx W        = "; .MaxVertexW
    Print #1, "Max Vol Extent  = "; .MaxVolumeExtent
    Print #1, "Pixel Shader Ver= "; Hex8(.PixelShaderVersion)
    Print #1, "Present Intervls= "; Hex8(.PresentationIntervals)
    Print #1, "Primitive Misc C= "; Hex8(.PrimitiveMiscCaps)
    Print #1, "Raster Caps     = "; Hex8(.RasterCaps)
    Print #1, "Shade Caps      = "; Hex8(.ShadeCaps)
    Print #1, "Src Blend Caps  = "; Hex8(.SrcBlendCaps)
    Print #1, "Stencil Caps    = "; Hex8(.StencilCaps)
    Print #1, "Texture AddrCaps= "; Hex8(.TextureAddressCaps)
    Print #1, "Texture Caps    = "; Hex8(.TextureCaps)
    Print #1, "Texture  OpCaps = "; Hex8(.TextureOpCaps)
    Print #1, "Vx Proc Caps    = "; Hex8(.VertexProcessingCaps)
    Print #1, "Vx Shader Vers  = "; Hex8(.VertexShaderVersion)
    Print #1, "Vol Tex Addr Cap= "; Hex8(.VolumeTextureAddressCaps)
    Print #1, "Vol Tex Filt Cap= "; Hex8(.VolumeTextureFilterCaps)
    Print #1, "Z Cmp Caps      = "; Hex8(.ZCmpCaps)
  End With
  Close #1

End Sub

Private Function Hex8(ByVal v As Long) As String

  Hex8 = "0x" & Right$("00000000" & Hex$(v), 8)

End Function

'Used to Conveniently Check if a Given Mode can Be selected
Private Function CheckDisplayMode(ByVal Width As Long, ByVal Height As Long, ByVal Depth As Long) As CONST_D3DFORMAT

 Dim i As Long
 Dim DispMode As D3DDISPLAYMODE

  '//1. Scan through
  For i = 0 To D3D8.GetAdapterModeCount(0) - 1
    Call D3D8.EnumAdapterModes(0, i, DispMode)
    If DispMode.Width = Width Then
      If DispMode.Height = Height Then
        If (DispMode.Format = D3DFMT_R5G6B5) _
           Or (DispMode.Format = D3DFMT_X1R5G5B5) _
           Or (DispMode.Format = D3DFMT_X4R4G4B4) Then       '16 bit mode
          If Depth = 16 Then
            CheckDisplayMode = DispMode.Format
            Exit Function
          End If
        ElseIf (DispMode.Format = D3DFMT_R8G8B8) _
               Or (DispMode.Format = D3DFMT_X8R8G8B8) Then   '32bit mode
          If Depth = 32 Then
            CheckDisplayMode = DispMode.Format
            Exit Function
          End If
        End If
      End If
    End If
  Next i
  CheckDisplayMode = D3DFMT_UNKNOWN

End Function

Private Sub GeneralErrorMsg(Error As ErrObject)

  MsgBox "DX Error: " & Error.Number, Error.Description, D3DX.GetErrorString(Error.Number)

End Sub

'-------------------------------- A very basic attempt at saving your work to disk --------------------------
'THIS DOESNT WORK AT ALL YET
Private Sub SavePic(ByRef Path As String, ByVal Format As CONST_D3DXIMAGE_FILEFORMAT)

 Dim Surface As Direct3DSurface8
 Dim SurfaceDesc As D3DSURFACE_DESC
 Dim Palette(0 To 255) As PALETTEENTRY
 Dim R As RECT

  On Error Resume Next
  Call D3DD.GetFrontBuffer(Surface)
  If Err.Number = D3D_OK Then
    MsgBox "got surface"
    Call Surface.GetDesc(SurfaceDesc)
    If Err.Number = D3D_OK Then
      MsgBox "got surface desc"
      R.bottom = SurfaceDesc.Height
      R.Right = SurfaceDesc.Width
      Call D3DD.GetPaletteEntries(0, Palette(0))
      If Err.Number = D3D_OK Then
        MsgBox "got palette"
        Call D3DX.SaveSurfaceToFile(Path, Format, Surface, Palette(0), R)
      End If
    End If
  End If

End Sub

'======================= a generic Has a Key Been HIT ==================================================
'given a list of keys return the keykode hit if it was in the list, else return -1
Public Function KeyHit(ParamArray vbKeys() As Variant) As Long
Attribute KeyHit.VB_Description = "Check wheter a key has been hit"

 Dim i As Long

  Call Keyboard.Acquire                                                       'in case it got lost
  Call Keyboard.GetDeviceStateKeyboard(KeyboardState)                         'get keyboard data
  For i = 0 To UBound(vbKeys)
    If KeyboardState.Key(CLng(mDIKeys.vbToDIK(vbKeys(i)))) Then               'see mapping in mDIKeys
      KeyHit = vbKeys(i)
      Exit Function
    End If
  Next i
  KeyHit = -1

End Function

'======================= a generic MouseMove ===========================================================

Public Sub GetMouseDelta(ByRef dXYZ As D3DVECTOR)
Attribute GetMouseDelta.VB_Description = "Get DirectX mouse movement Values"

  Call Mouse.Acquire                                                          'in case it got lost
  Call Mouse.GetDeviceStateMouse(MouseState)                                  'get mouse data
  dXYZ.x = MouseState.lX
  dXYZ.y = MouseState.lY
  dXYZ.z = MouseState.lZ

End Sub

'======================= a generic Play WAV function ==================================================

Public Sub PlaySoundFromFile(ByRef Path As String, Optional ByVal LoopIt As Boolean = False)
Attribute PlaySoundFromFile.VB_Description = "Play the named WAV file, optionally looping"

'using the global variables , make a buffer (previous one is blown away)

  If Not Sound Is Nothing Then                                           'we may already be playing something
    If (Sound.GetStatus And DSBSTATUS_PLAYING) = DSBSTATUS_PLAYING Then
      Sound.Stop                                                         'Kill that sound
    End If
    Set Sound = Nothing
  End If

  If Len(Path) <> 0 Then                                         'If we pass an empty path we do nothing
    On Error Resume Next
    Set Sound = DS8.CreateSoundBufferFromFile(Path, SoundDesc)   'get it from File
    If Sound Is Nothing Then
      MsgBox "DirectSound cannot load:" & vbCrLf & Path & vbCrLf _
             & "This program will carry on without it", vbInformation
    Else
      If LoopIt Then
        Sound.Play DSBPLAY_LOOPING            'play as loop
      Else
        Sound.Play DSBPLAY_DEFAULT            'play once
      End If
    End If
  End If

End Sub

Private Sub Class_Initialize()

  Call mDIKeys.MakevbToDIK

End Sub

Private Sub Class_Terminate()

  Call ShutDownDX8

End Sub

':) Ulli's VB Code Formatter V2.13.5 (01-Feb-03 20:53:11) 115 + 411 = 526 Lines
