VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cScene"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'RVTDX.cScene.cls - a component of the rvtDX.dll
'©2003 Ron van Tilburg - rivit@f1.net.au
'Freeware for Educational Purposes, For commercial interests contact author please, I retain copyright.

'cScene.cls   The Master of Ceremonies -  All rendering Done through Here

'The Scene Maintains a small Set of Vertex and Index Buffers for the entire Render
'We reuse Figures if they are the same
'We have separate FigureBuffers for PointLists and Linelists or TriangleLists
'Nothing is allocated until it is needed, and the largest size of the Buffers is constrained
'to the DevCaps limit

Private Type FigureTableEntry
  FigSpec       As FIG_SPEC
  Primitive     As CONST_D3DPRIMITIVETYPE   'Primitive Type
  VxBufferIx    As Integer                  'Which VxBuffer
  FirstVxEntry  As Long                     'Where To start
  NVxEntries    As Long                     'Nr of Vx Entries to Use
  IxBufferIx    As Integer                  'Which IxBuffer
  FirstIxEntry  As Long                     'Where To start
  NIxEntries    As Long                     'Nr of Ix Entries to Use
  NPrimitives   As Long                     'The Nr Of Primitives it Contains
End Type

Private Type VertexBufferEntry              'USED when rendering a Figure
  VxBuffer      As Direct3DVertexBuffer8    'The Vertex Buffer
  VxFormat      As Long                     'Its Format         VertexFMT for now
  VxSize        As Long                     'and Size           Len(ZVertex) for now
  NVx           As Long                     'Nr of Entries      DevCaps.MaxVertex
  VxFirstFree  As Long                      'Next one to Use
End Type

Private Type IndexBufferEntry               'USED when rendering a Figure
  IxBuffer      As Direct3DIndexBuffer8     'The Index Buffer
  IxSize        As Long                     'Size of Entries    16 for now
  NIx           As Long                     'Nr of Entries      DevCaps.MaxVxIndex
  IxFirstFree   As Long                     'Next one to Use
End Type

'==================================== INSTANCE DATA =========================================================
Private zRenderFlags  As RENDER_FLAGS       'Will Override Scene RenderFlags
Private zPrepared     As Boolean            'True When Renderstate=RenderFlags

Private zLengthScale  As Single             'The Scaling Factor from Real Life to Scene Length
Private zFrameRate    As Single             'The desired FrameRate

Private zAmbient      As Long               'The Scene Ambient Colour     (default &HFFFFFFFF)
Private zBGColour     As Long               'The Scene Background Colour  (default &H00000000)
Private zFont         As D3DXFont           'Font For General Use
Private zFontHUD      As D3DXFont           'Font For HUD
Private zShowHUD      As Boolean            'True if We Do
Private zTitle        As String             'Scene Title

Private zLights()     As cLight             'The Lights, zLights(0)=Ambient Light
Private zNLights      As Integer            'the number of figures

Private zCamera       As cFigure            'What we are looking With
Private zViewpoint    As cFigure            'The Spot where we look at
Private zAxes         As Long               'The FigureSpec Index Of the Axes (if -1) there are none

Private zFigures()    As cFigure            'The list of figures
Private zNFigures     As Integer            'the number of figures

Private zTextures()   As Direct3DTexture8   'All of the textures we hope to use
Private zNTextures    As Long

Private zDefMat       As D3DMATERIAL8       'A default Material (White in every sense)
'-------------------------------------------' render specific data

Private zFigSpecs()   As FigureTableEntry   'The Unique Specs we have assembled so far
Private zNFigSpecs    As Integer            'How many of them

Private zVxBuffers()  As VertexBufferEntry  'The Tables of Stored Vertex Buffers
Private zNVxBuffers   As Integer            'How Many So Far

Private zIxBuffers()  As IndexBufferEntry   'The Tables of Stored Index Buffers
Private zNIxBuffers   As Integer            'How Many So Far

'-------------------------------------------
Private zNFrames      As Long               'No of Frames played since Last Reset
Private zFrameTimer   As Long               'A HighRes Elapsed Timer

'=========================== INIT AND TERMINATION ===========================================================

Private Sub Class_Initialize()

  ReDim zFigures(0 To 9)                    'default 9
  Dim AFont As StdFont

  Call SetScaling                           '1 metre in camera space = 1 metre
  '1 Increment = 1 Frame

  Set zViewpoint = New cFigure              'make the camera viewpoint (default 0,0,0)
  zViewpoint.FigSpec = FIGS_LOCATOR

  Set zCamera = New cFigure                 'make the camera
  zCamera.FigSpec = FIGS_LOCATOR
  Call MoveCameraRAE(120, -36, 36)          'The Camera is automatically made and nominally positioned

  zAxes = -1
  Call SetProjection
  Call SetAmbientLight(192, 192, 192)

  With zDefMat
    With .Ambient
      .a = 1: .b = 1: .g = 1: .R = 1
    End With
    .Diffuse = .Ambient
    .specular = .Ambient
    .Power = 1
  End With

  Set AFont = New StdFont
  AFont.Name = "Arial"
  AFont.Size = 8
  Set zFontHUD = CreateDXFont(AFont)
  AFont.Size = 10
  Set zFont = CreateDXFont(AFont)
  
  'Prepare for FPS
  zNFrames = -1
  zFrameTimer = StopWatchInitialize()

  Call UpdateRenderState(0, -1, Reset:=True) 'everything to our defaults
End Sub

Private Sub Class_Terminate()       'free up all attached Figures

 Dim i As Long

  For i = 0 To zNVxBuffers - 1              'Kill VertexBuffers
    Set zVxBuffers(i).VxBuffer = Nothing
  Next i

  For i = 0 To zNIxBuffers - 1              'Kill IndexBuffers
    Set zIxBuffers(i).IxBuffer = Nothing
  Next i

  For i = 0 To zNFigures - 1                'Kill All Figures
    Set zFigures(i) = Nothing
  Next i

  Set zCamera = Nothing
  Set zViewpoint = Nothing

  For i = 0 To zNLights - 1                 'Kill All Lights
    Set zLights(i) = Nothing
  Next i

  For i = 0 To 7                            'Clear all textures in Render Path
    Call D3DD.SetTexture(i, Nothing)
  Next i

  For i = 0 To zNTextures - 1               'Kill off the Textures
    Set zTextures(i) = Nothing
  Next i

  Set zFont = Nothing                       'free the USR font
  Set zFontHUD = Nothing                    'free the HUD font
  
  Call D3DD.ResourceManagerDiscardBytes(0)  'Free Anything It Can

End Sub

'=========================================================================================================
'============================= SOME BASIC SCENE ELEMENTS =================================================
'=========================================================================================================

'This doesnt do anything yet - for future autoscale and FrameRate control
Public Sub SetScaling(Optional ByVal LengthScale As Single = 1, Optional ByVal FrameRate As Single = 30)
Attribute SetScaling.VB_Description = "Set Length and Framerate scales desired (not used yet)"

  If LengthScale <= 0 Then LengthScale = 1
  If FrameRate <= 0 Then FrameRate = 30
  zLengthScale = LengthScale                  'Affects Scalings - used to Reduce to Unit Cube
  zFrameRate = FrameRate                      'Affects speeds of animation

End Sub

Public Property Get Camera() As cFigure
Attribute Camera.VB_Description = "Returns/Sets the Camera Figure"

  Set Camera = zCamera

End Property

Public Property Get Viewpoint() As cFigure
Attribute Viewpoint.VB_Description = "Set/Return Viewpoint info"

  Set Viewpoint = zViewpoint

End Property

'==========================================================================================================
'===================================== AXES & HUD =========================================================
'==========================================================================================================

Public Sub AddAxes(ByVal AxisType As FIG_SPEC)
Attribute AddAxes.VB_Description = "Show AXES or AXISPLANES in the scene"

 Dim Axes As cFigure

  If AxisType = FIGS_AXES Or AxisType = FIGS_AXISPLANES Then
    Set Axes = New cFigure                      'make the axes, just another figure == actually zFigures(0)
    Axes.FigSpec = AxisType
    Call Axes.SetScale(100, 100, 100)           'a default scale
    Axes.RenderFlags = RF_VERTEXCOLOURS         'always in Vertex Colours
    Call AddFigure(Axes)
    zAxes = zNFigures - 1
  End If

End Sub

Public Property Get Axes() As cFigure
Attribute Axes.VB_Description = "Returns/Sets the AXIS figure"

  If zAxes <> -1 Then
    Set Axes = zFigures(zAxes)
  End If

End Property

Public Property Let ShowHUD(ByVal TrueFalse As Boolean)
Attribute ShowHUD.VB_Description = "Turn the Heads up display on or Off"

  zShowHUD = TrueFalse

End Property

Public Property Let Title(ByRef NewTitle As String)
Attribute Title.VB_Description = "Set a title for the Scene"
  zTitle = NewTitle
End Property

'==========================================================================================================
'===================================== FONTS AND TEXT =====================================================
'==========================================================================================================
Property Set Font(ByRef NewFont As StdFont)     'Default Arial 10
Attribute Font.VB_Description = "Set the Font for Titling"

  Dim dxfont As D3DXFont
  
  Set dxfont = CreateDXFont(NewFont)
  If Not dxfont Is Nothing Then
    Set zFont = dxfont
  End If

End Property

'Draw Text in Usr Font
Private Sub DrawText(ByVal XLeft As Long, ByVal YTop As Long, _
                    ByVal XRight As Long, ByVal YBottom As Long, _
                    ByRef Text As String, _
                    Optional ByVal ColourRGB As Long = &HFFFFFFFF, _
                    Optional ByVal FormatFlags As Long = 0)
  Dim R As RECT
  R.Left = XLeft
  R.Top = YTop
  R.Right = XRight
  R.bottom = YBottom
  Call zFont.DrawTextW(Text, Len(Text), R, FormatFlags, ColourRGB Or &HFF000000)
End Sub

'==========================================================================================================
'========================== PROJECTIONS ==== ONLY PERSPECTIVE FOR NOW =====================================
'==========================================================================================================

Public Sub SetProjection(Optional ByVal ViewField As Single = 60, _
                         Optional ByVal FrontPlane As Single = 0.1, _
                         Optional ByVal BackPlane As Single = 100000)
Attribute SetProjection.VB_Description = "Set the depth and with of the PERSPECTIVE PROJECTION we are using"

 Dim tmpMx As D3DMATRIX, t As Single

  If BackPlane < FrontPlane Then
    t = BackPlane: BackPlane = FrontPlane: FrontPlane = t
  End If
  ViewField = DtoR * Abs(ViewField) / 2
  If FrontPlane < 0.1 Then FrontPlane = 0.1
  If BackPlane < 0.1 Then BackPlane = 100000

  Call D3DXMatrixPerspectiveFovLH(tmpMx, ViewField, D3DDM.Height / D3DDM.Width, FrontPlane, BackPlane)
  Call D3DD.SetTransform(D3DTS_PROJECTION, tmpMx)                              'Establish Projection

End Sub

'==========================================================================================================
'============================== TEXTURE POOL =============================================================
'==========================================================================================================

Public Sub TextureFromFile(ByRef Path As String)
Attribute TextureFromFile.VB_Description = "Load a texture into the scene from File"

 Dim Texture As Direct3DTexture8

  If Len(Path) <> 0 Then
    On Error Resume Next
    Set Texture = D3DX.CreateTextureFromFile(D3DD, Path)
    On Error GoTo 0
  End If

  If Not Texture Is Nothing Then
    ReDim Preserve zTextures(0 To zNTextures)
    Set zTextures(zNTextures) = Texture
    zNTextures = zNTextures + 1
  End If

End Sub

'==========================================================================================================
'============================= LIGHTS & LIGHTING ============================================================
'==========================================================================================================
Public Sub SetBackground(ByVal Red As Integer, Green As Integer, Blue As Integer)
Attribute SetBackground.VB_Description = "Set the colour of the back of the scene &HFFRRGGBB"

  zBGColour = &HFF000000 Or ((Red And &HFF) * 65536 + (Green And &HFF) * 256& + (Blue And &HFF))

End Sub

Public Sub SetAmbientLight(ByVal Red As Long, ByVal Green As Long, ByVal Blue As Long)
Attribute SetAmbientLight.VB_Description = "Set the value of background light &HFFRRGGBB"

  zAmbient = &HFF000000 Or (Red And &HFF) * 65536 + (Green And &HFF) * 256& + (Blue And &HFF)
  Call D3DD.SetRenderState(D3DRS_AMBIENT, zAmbient)

End Sub

Public Sub AddLight(ByRef NewLight As cLight)
Attribute AddLight.VB_Description = "Add a defined light into the scene"

  ReDim Preserve zLights(0 To zNLights)
  Set zLights(zNLights) = NewLight
  With zLights(zNLights)
    Call .AddMe(zNLights)
    .State = LIGHT_ON
  End With
  zNLights = zNLights + 1

End Sub

Public Property Get Light(ByVal Index As Long) As cLight
Attribute Light.VB_Description = "Data about the nth light"

  If Index >= 0 And Index < zNLights Then
    Set Light = zLights(Index)
  End If

End Property

Public Sub SwitchLight(ByVal Index As Long, ByVal NewState As LIGHTS_ONOFF)
Attribute SwitchLight.VB_Description = "Turn a light on or Off"

  If Index >= 0 And Index < zNLights Then
    zLights(Index).State = NewState               'Flip The switch
  End If

End Sub

'==========================================================================================================
'============================= FIGURES =====================================================================
'==========================================================================================================

Public Property Get Figure(ByVal FigureNr As Long) As cFigure
Attribute Figure.VB_Description = "Return the Object of the ith figure"

  If FigureNr >= 0 And FigureNr < zNFigures Then
    Set Figure = zFigures(FigureNr)
  End If

End Property

'When a Figure is Added its VertexBuffers are generated. If CUSTOM by the Generator, else by Template
'Templated Figures are only created once and shared as many times as needed
'They are Indexed by FigSpec,  if not already there they are generated
'return as True indicates success

Public Function AddFigure(ByRef Figure As cFigure, _
                          Optional ByRef Generator As Object) As Boolean
Attribute AddFigure.VB_Description = "Add a defined figure to the scene"

 Dim i As Long, FS As Long

  'We first look to see if we already have a figure exactly the same as this one
  FS = Figure.FigSpec
  
  If (FS And FIGS_TYPEMASK) = FIGS_MESH Then                                'We need to generate a Mesh
    AddFigure = Figure.GenerateMesh
  Else                                                                      'We need to Generate Vertices
    For i = 0 To zNFigSpecs - 1
      If zFigSpecs(i).FigSpec = FS Then Exit For         'Its already there
    Next i
  
    If i = zNFigSpecs Then                                                  'we need a new one
      ReDim Preserve zFigSpecs(0 To zNFigSpecs)
      zNFigSpecs = zNFigSpecs + 1                                           'we have one more of them
      If (FS And FIGS_CUSTOM) = FIGS_CUSTOM Then                            'Independent Generator needed
        If Generator Is Nothing Then Exit Function                          'Cant find Generator
        If Generator.Generate(FS, zVx(), zIx(), Primitive) Then             'Pass Pointers
          AddFigure = BuildVIBuffers(i, FS)                                 'and its FigureBuffer is Now Filled
        End If
      Else                                                                  'We may use a stock Figure
        If GenerateFigure(FS) Then AddFigure = BuildVIBuffers(i, FS)
      End If
    End If
  
    If zFigSpecs(i).FigSpec = FS Then
      AddFigure = True
      Figure.FigSpecIndex = i                                               'set it as the Figure's Buffer
    End If
    Call EraseWorkArrays    'Clear the zVx and zIx  Generator Arrays
  End If
  
  If AddFigure Then
    Set zFigures(zNFigures) = Figure
    zNFigures = zNFigures + 1
    If zNFigures > UBound(zFigures) Then
      ReDim Preserve zFigures(0 To zNFigures + 9)
    End If
  End If

End Function

'This rather Tricky routine maintains the Vertex and Index Buffers which we are attempting to keep
'down in number. For every FIGSPEC we check if it will fit in the buffers and if not make new ones
'All of the indexing info is kept in zFigSpecs(i)

Private Function BuildVIBuffers(ByVal i As Long, ByVal FigSpec As FIG_SPEC) As Boolean    'TRUE is OK

 Dim NVx As Long, NIx As Long, MaxSize As Long
 Dim NewBufReqd As Boolean, NewFigSpec As FigureTableEntry

  NVx = 1 + UBound(zVx) 'A Global() set by Generator
  NIx = 1 + UBound(zIx) 'A Global() set by Generator

  MaxSize = D3DDC.MaxPrimitiveCount
  If MaxSize > 65535 Then MaxSize = 65535
  If NVx < 1 Then Exit Function                         'nothing to add
  If NVx > MaxSize Then Exit Function                   'we cant make a buffer this large

  If Primitive <> D3DPT_POINTLIST Then                  'No Index Entries for POINTLISTS
    MaxSize = D3DDC.MaxVertexIndex
    If MaxSize > 65535 Then MaxSize = 65535
    If NIx < 1 Then Exit Function                       'nothing to add
    If NIx > MaxSize Then Exit Function                 'we cant make a buffer this large
  End If

  With NewFigSpec
    .FigSpec = FigSpec
    .Primitive = Primitive                              'A Global set by Generator
    If Primitive <> D3DPT_POINTLIST Then                'No Index Entries for POINTLISTS
      .NPrimitives = NIx \ PrimitiveNVx(Primitive)
    Else
      .NPrimitives = NVx
    End If
  End With

  'CheckIf We Need to Make a Vertex Buffer
  If zNVxBuffers = 0 Then
    NewBufReqd = True
  Else
    With zVxBuffers(zNVxBuffers - 1)
      If .NVx - NVx < .VxFirstFree Then NewBufReqd = True
    End With
  End If

  If NewBufReqd Then
    If Not MakeVertexBuffer(LenB(zVx(0)), VertexFMT) Then Exit Function     'Failed so give up
  End If

  NewBufReqd = False                                              'We now have have a buffer large enough

  With zVxBuffers(zNVxBuffers - 1)
    If D3DVertexBuffer8SetData(.VxBuffer, _
       .VxFirstFree * .VxSize, _
       NVx * .VxSize, _
       0, _
       zVx(0)) <> D3D_OK Then Exit Function ' Copy the vertices the buffer

    NewFigSpec.VxBufferIx = zNVxBuffers - 1                         'This VertexBuffer
    NewFigSpec.FirstVxEntry = .VxFirstFree                          'Starting Here
    NewFigSpec.NVxEntries = NVx                                     'For N

    .VxFirstFree = .VxFirstFree + NVx                               ' Mark next spot to use
  End With

  'Now Deal with Index Buffers (these are NOT used for POINTLISTS)
  If Primitive = D3DPT_POINTLIST Then
    NewFigSpec.IxBufferIx = -1                                      'This nonexistent VertexBuffer
    NewFigSpec.FirstIxEntry = -1                                    'Starting Here ??
    NewFigSpec.NIxEntries = 0                                       'For N
  Else
    'CheckIf We Need to Make a Vertex Buffer
    If zNIxBuffers = 0 Then
      NewBufReqd = True
    Else
      With zIxBuffers(zNIxBuffers - 1)
        If .NIx - NIx < .IxFirstFree Then NewBufReqd = True
      End With
    End If

    If NewBufReqd Then
      If Not MakeIndexBuffer(LenB(zIx(0)), D3DFMT_INDEX16) Then Exit Function     'Failed so give up
    End If

    NewBufReqd = False                                              'We now have have a buffer large enough

    With zIxBuffers(zNIxBuffers - 1)
      If D3DIndexBuffer8SetData(.IxBuffer, _
         .IxFirstFree * .IxSize, _
         NIx * .IxSize, _
         0, _
         zIx(0)) <> D3D_OK Then Exit Function ' Copy the vertices the buffer

      NewFigSpec.IxBufferIx = zNIxBuffers - 1                         'This IndexBuffer
      NewFigSpec.FirstIxEntry = .IxFirstFree                          'Starting Here
      NewFigSpec.NIxEntries = NIx                                     'For N

      .IxFirstFree = .IxFirstFree + NIx                               ' Mark next spot to use
    End With
  End If

  'So having Got This Far we can Add the New FigSpecEntry
  zFigSpecs(i) = NewFigSpec
  BuildVIBuffers = True         'YIPPEEEEEE

End Function

'We always Make a New One at the End (as large as possible)
Private Function MakeVertexBuffer(ByVal VxSize As Long, ByVal VertexFVF As Long) As Boolean

 Dim VxBuffer As Direct3DVertexBuffer8, MaxSize As Long

  MaxSize = D3DDC.MaxPrimitiveCount
  If MaxSize > 65535 Then MaxSize = 65535           'a 2MEG allocation
  Set VxBuffer = D3DD.CreateVertexBuffer(MaxSize * VxSize, 0, VertexFVF, D3DPOOL_MANAGED)
  If VxBuffer Is Nothing Then Exit Function    'Failed at the first Hurdle

  ReDim Preserve zVxBuffers(0 To zNVxBuffers)
  With zVxBuffers(zNVxBuffers)
    Set .VxBuffer = VxBuffer
    .NVx = MaxSize
    .VxFirstFree = 0
    .VxFormat = VertexFVF
    .VxSize = VxSize
  End With
  zNVxBuffers = zNVxBuffers + 1
  MakeVertexBuffer = True

End Function

'We always Make a New One at the End (as large as possible)
Private Function MakeIndexBuffer(ByVal IxSize As Long, ByVal IndexFMT As CONST_D3DFORMAT) As Boolean

 Dim IxBuffer As Direct3DIndexBuffer8, MaxSize As Long

  MaxSize = D3DDC.MaxVertexIndex
  If MaxSize > 65535 Then MaxSize = 65535               'a 128K allocation
  Set IxBuffer = D3DD.CreateIndexBuffer(MaxSize * IxSize, 0, IndexFMT, D3DPOOL_MANAGED)
  If IxBuffer Is Nothing Then Exit Function                           'Failed at the Second Hurdle

  ReDim Preserve zIxBuffers(0 To zNIxBuffers)
  With zIxBuffers(zNIxBuffers)                  ' Create a vertex buffer with enough memory for type ZVertex
    Set .IxBuffer = IxBuffer
    .IxFirstFree = 0
    .IxSize = IxSize
    .NIx = MaxSize
  End With
  zNIxBuffers = zNIxBuffers + 1
  MakeIndexBuffer = True

End Function

Private Function PrimitiveNVx(ByVal Primitive As CONST_D3DPRIMITIVETYPE) As Long

'only these three used

  Select Case Primitive:
  Case D3DPT_POINTLIST:     PrimitiveNVx = 1
  Case D3DPT_LINELIST:      PrimitiveNVx = 2
  Case D3DPT_TRIANGLELIST:  PrimitiveNVx = 3
  End Select

End Function

'==========================================================================================================
'============================= THE RENDER STAGE ===========================================================
'==========================================================================================================

Public Property Let RenderFlags(ByVal NewFlags As RENDER_FLAGS)
Attribute RenderFlags.VB_Description = "The parameters for Display of the Entire Scene. Can be ovverridden at Figure Level"
  zRenderFlags = NewFlags
  zPrepared = False
End Property

Public Property Get RenderFlags() As RENDER_FLAGS
  RenderFlags = zRenderFlags
End Property

'---------------------------------------------------------------------------------------------------------
'Used to Keep control of render settings, it returns the Settings in Effect
Private Function UpdateRenderState(ByVal CurrentState As RENDER_FLAGS, _
                                   ByVal RenderFlags As RENDER_FLAGS, _
                                   Optional Reset As Boolean = False) As RENDER_FLAGS  'applies all of the default settings based on renderflags
    
  With D3DD
    If Reset Then                                     'Put back The bits passed in
      If RenderFlags <> 0 Then
        If (RenderFlags And RF_UNLIT) <> 0 Then
          Call .SetRenderState(D3DRS_LIGHTING, 1)
        End If
        
        If (RenderFlags And RF_VERTEXCOLOURS) <> 0 Then
          Call .SetRenderState(D3DRS_DIFFUSEMATERIALSOURCE, D3DMCS_MATERIAL)    'use material vertex diffuse
        End If
        
        If (RenderFlags And RF_DRAWPOINTS) <> 0 _
        Or (RenderFlags And RF_DRAWWIREFRAME) <> 0 Then
          Call .SetRenderState(D3DRS_FILLMODE, D3DFILL_SOLID)       ' render solid polyhedra
        End If
        
        If (RenderFlags And RF_REMOVEHIDDEN) <> 0 Then
          Call .SetRenderState(D3DRS_CULLMODE, D3DCULL_NONE)      ' keep everything
        End If
        
        If (RenderFlags And RF_SHADEFLAT) <> 0 Then
          Call .SetRenderState(D3DRS_SHADEMODE, D3DSHADE_GOURAUD) ' Use Gouraud Shading
        End If
        
        If (RenderFlags And RF_SHINY) <> 0 Then
          Call .SetRenderState(D3DRS_SPECULARENABLE, 0)
        End If
        
        If (RenderFlags And RF_ALPHAOPS) <> 0 Then
          Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_ONE)
          Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_ZERO)
          Call .SetRenderState(D3DRS_ALPHABLENDENABLE, 0)
        End If
      End If
      UpdateRenderState = CurrentState And (Not RenderFlags)
    Else
      If RenderFlags <> CurrentState Then                        ' set the bits passed in
        If (RenderFlags And RF_UNLIT) <> 0 Then
          Call .SetRenderState(D3DRS_LIGHTING, 0)
        End If
        
        If (RenderFlags And RF_VERTEXCOLOURS) <> 0 Then
          Call .SetRenderState(D3DRS_DIFFUSEMATERIALSOURCE, D3DMCS_COLOR1)    'use vertex diffuse
        End If
        
        If (RenderFlags And RF_DRAWPOINTS) <> 0 Then
          Call .SetRenderState(D3DRS_FILLMODE, D3DFILL_POINT)     ' render solid points
        ElseIf (RenderFlags And RF_DRAWWIREFRAME) <> 0 Then
          Call .SetRenderState(D3DRS_FILLMODE, D3DFILL_WIREFRAME) ' render solid wireframe
        Else
          Call .SetRenderState(D3DRS_FILLMODE, D3DFILL_SOLID)     ' render solid polyhedra
        End If
        
        If (RenderFlags And RF_REMOVEHIDDEN) <> 0 Then
          Call .SetRenderState(D3DRS_CULLMODE, D3DCULL_CCW)       ' remove all counterclockwise faces
        End If
        
        If (RenderFlags And RF_SHADEFLAT) <> 0 Then
          Call .SetRenderState(D3DRS_SHADEMODE, D3DSHADE_FLAT)    ' Use Flat Shading
        End If
        
        If (RenderFlags And RF_SHINY) <> 0 Then
          Call .SetRenderState(D3DRS_SPECULARENABLE, 1)
        End If
        
        If (RenderFlags And RF_LIGHTTINT) <> 0 Then
          Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_ONE)
          Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_DESTCOLOR)
          Call .SetRenderState(D3DRS_ALPHABLENDENABLE, 1)
        End If
        
        If (RenderFlags And RF_DARKTINT) <> 0 Then
          Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCCOLOR)
          Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVSRCCOLOR)
          Call .SetRenderState(D3DRS_ALPHABLENDENABLE, 1)
        End If
        
        If (RenderFlags And RF_TRANSPARENT) <> 0 Then
          Call .SetRenderState(D3DRS_SRCBLEND, D3DBLEND_SRCALPHA)
          Call .SetRenderState(D3DRS_DESTBLEND, D3DBLEND_INVDESTCOLOR)
          Call .SetRenderState(D3DRS_ALPHABLENDENABLE, 1)
        End If
        
      End If
      UpdateRenderState = CurrentState Or RenderFlags
    End If
  End With
End Function

'---------------------------------------------------------------------------------------------------------
Public Function RenderReset() As RENDER_FLAGS
Attribute RenderReset.VB_Description = "Force reset Render State"
  RenderReset = UpdateRenderState(0, -1, Reset:=True)
  RenderReset = UpdateRenderState(RenderReset, zRenderFlags)
  Call ResetFPS
  zPrepared = True
End Function

'---------------------------------------------------------------------------------------------------------

Public Sub Render(Optional AutoIncrement As Boolean = False)
Attribute Render.VB_Description = "Render a complete frame of the Scene, optionally automatically incrementing everything"

 Dim i As Long, FigI As cFigure, currVxBuffer As Integer, tmpMat  As D3DMATRIX
 Dim RenderState As Long
  
  If Not zPrepared Then Call RenderReset       'done whenever scene renderflags have been changed
  RenderState = zRenderFlags
  
  'If zNFrames > 200 And GetFPS > zFrameRate Then Exit Sub       'if we are going to fast loop back now
  
  'Establish WorldView (Projection is usually Done Once before RenderLoop)
  Call D3DXMatrixLookAtLH(tmpMat, zCamera.CurrWorldPos, zViewpoint.CurrWorldPos, WORLD_YISUP)
  Call D3DD.SetTransform(D3DTS_VIEW, tmpMat)                                    'Establish World View

  ' clear the rendering device backbuffer and z-buffer
  Call D3DD.Clear(0, ByVal 0, D3DCLEAR_TARGET Or D3DCLEAR_ZBUFFER, zBGColour, 1#, 0)
  On Error Resume Next
  Call D3DD.BeginScene              ' start rendering

  currVxBuffer = -1                 'Mark as nothing connected to Renderstream yet
  For i = 0 To zNFigures - 1        'Work our way through all the figures in the scene
    Set FigI = zFigures(i)
    If Not FigI Is Nothing Then     'paranoia
      If (FigI.FigSpec And FIGS_TYPEMASK) <> FIGS_LOCATOR Then            'No rendering for FIGS_LOCATOR
        If (FigI.RenderFlags And RF_IGNORE) = 0 Then                      'No rendering for RF_IGNORE

          RenderState = UpdateRenderState(zRenderFlags, zRenderFlags Xor FigI.RenderFlags)  'Fix Modes
          
          Call DoMaterialSetup(FigI, RenderState)                         'Deal with The Materials
          Call DoTextureSetup(FigI, RenderState)                          'Deal with The textures
          
          With FigI
            Call D3DD.SetTransform(D3DTS_WORLD, .GetWorldMatrix())        'Set the world transform
            
            If (.FigSpec And FIGS_TYPEMASK) = FIGS_MESH Then
              If Not .Mesh Is Nothing Then                                'We will Draw a DX Mesh
                Call .Mesh.DrawSubset(0)                                  'all of it
                currVxBuffer = -1                                         'we need to set it again
              End If
            Else
              With zFigSpecs(.FigSpecIndex)                               'We will Draw Our Vertices
                'Check Vertex Buffers
                If .VxBufferIx <> currVxBuffer Then                       'We Need to Change VertexBuffer
                  With zVxBuffers(.VxBufferIx)
                    Call D3DD.SetVertexShader(.VxFormat)                  'Set the vertex format
                    Call D3DD.SetStreamSource(0, .VxBuffer, .VxSize)      'render from it
                  End With
                End If
                currVxBuffer = .VxBufferIx
    
                If .Primitive = D3DPT_POINTLIST Then
                  Call D3DD.DrawPrimitive(D3DPT_POINTLIST, .FirstVxEntry, .NPrimitives) 'And Draw It
                Else
                  Call D3DD.SetIndices(zIxBuffers(.IxBufferIx).IxBuffer, .FirstVxEntry) 'using this set of indices
                  Call D3DD.DrawIndexedPrimitive(.Primitive, 0, .NIxEntries, .FirstIxEntry, .NPrimitives)                                               'And Draw It
                End If
              End With
            End If
          End With
          Call UpdateRenderState(RenderState, zRenderFlags Xor RenderState, Reset:=True)  'ResetModes
        End If
      End If
    End If
  Next i

  If zShowHUD Then Call DrawHUD                     ' Draw Heads Up Display
  If Len(zTitle) <> 0 Then Call DrawTitle           ' Draw Scene Title
  
  If (zNFrames Mod 5) = 0 Then DoEvents             ' Give Windows a chance too (done here to prevent screen flash)
  
  On Error GoTo 0
  Call D3DD.EndScene                                ' End rendering
  Call D3DD.Present(ByVal 0, ByVal 0, 0, ByVal 0)   ' Present the contents by flipping it to the screen
                                                    ' NOTE THIS CAN FAIL IF THE SCREEN HAS LOST FOCUS
                                                    ' I HAVE YET TO DEAL WITH THIS CIRCUMSTANCE
  
  If AutoIncrement Then Call Increment              ' Figures, Camera, ViewPoint and Lights can Move
  Call IncrementFramecount                          ' New Framecount

End Sub

Private Sub DoMaterialSetup(ByRef FigI As cFigure, ByVal RenderState As RENDER_FLAGS) 'MATERIAL HANDLING
  With FigI
    If (RenderState And RF_VERTEXCOLOURS) <> 0 Then
      Call D3DD.SetMaterial(zDefMat)                                          'Just Pretend
    Else
      If .MatMode <> FIGM_VXC Then                                            'set Materials if relevant
        Call D3DD.SetMaterial(.Material)
      Else
        Call D3DD.SetMaterial(zDefMat)                                        'Just Pretend
      End If
    End If
  End With
End Sub

Private Sub DoTextureSetup(ByRef FigI As cFigure, ByRef RenderState As RENDER_FLAGS)  'TEXTURE HANDLING
  Dim j As Long
  
  'Assume No Textures but if Scene or Figure require Textures then set them up
  With FigI
    Call D3DD.SetTexture(0, Nothing)                                         'Assume No Textures
    Call D3DD.SetTextureStageState(0, D3DTSS_COLOROP, D3DTOP_DISABLE)
    Call D3DD.SetTextureStageState(0, D3DTSS_ALPHAOP, D3DTOP_DISABLE)
    
    If (zRenderFlags And RF_NOTEXTURES) = 0 _
    And (.RenderFlags And RF_NOTEXTURES) = 0 Then                          'Set/Clear Textures
      For j = 0 To 7
        If .TextureIndex(j) = -1 _
           Or .TextureIndex(j) >= zNTextures Then    'It hasnt been Assigned (to Scene)
          Call D3DD.SetTexture(j, Nothing)
          Call D3DD.SetTextureStageState(j, D3DTSS_COLOROP, D3DTOP_DISABLE)   'disable stage j to 7
          Call D3DD.SetTextureStageState(j, D3DTSS_ALPHAOP, D3DTOP_DISABLE)   'disable stage j to 7
          Exit For
        Else
          Call D3DD.SetTexture(j, zTextures(.TextureIndex(j)))                'set texture
          Call D3DD.SetTextureStageState(j, D3DTSS_COLOROP, D3DTOP_MODULATE)  'enable colorops
          Call D3DD.SetTextureStageState(j, D3DTSS_COLORARG1, D3DTA_TEXTURE)
          Call D3DD.SetTextureStageState(j, D3DTSS_COLORARG2, D3DTA_CURRENT)
          If (.TextureRenderMode(j) And RF_ALPHAOPS) <> 0 Then
            RenderState = RenderState Or RF_ALPHAOPS                          'Will be reset Post Render
            Call D3DD.SetRenderState(D3DRS_ALPHABLENDENABLE, 1)
            Call D3DD.SetTextureStageState(j, D3DTSS_ALPHAOP, D3DTOP_BLENDCURRENTALPHA)
            Call D3DD.SetTextureStageState(j, D3DTSS_ALPHAARG1, D3DTA_TEXTURE)
            Call D3DD.SetTextureStageState(j, D3DTSS_ALPHAARG2, D3DTA_CURRENT)
            If (.TextureRenderMode(j) And RF_DARKTINT) <> 0 Then
              Call D3DD.SetTextureStageState(j, D3DTSS_ALPHAOP, D3DTOP_BLENDDIFFUSEALPHA)
            ElseIf (.TextureRenderMode(j) And RF_LIGHTTINT) <> 0 Then
              Call D3DD.SetTextureStageState(j, D3DTSS_ALPHAOP, D3DTOP_BLENDTEXTUREALPHA)
            ElseIf (.TextureRenderMode(j) And RF_TRANSPARENT) <> 0 Then
              Call D3DD.SetTextureStageState(j, D3DTSS_ALPHAOP, D3DTOP_ADDSMOOTH)
            End If
          End If
        End If
      Next j
    End If
  End With
End Sub

Private Sub DrawHUD()
  
  Dim v As D3DVECTOR, w As D3DVECTOR
  Dim rr As Single, Az As Single, el As Single
  
  v = zCamera.CurrWorldPos
  w = zViewpoint.CurrWorldPos
  Call RecttoSph(rr, Az, el, v.x - w.x, v.y - w.y, v.z - w.z)
  
  Call DrawTextHUD(10, 10, 300, 30, "VPT: (" & FmtVal(w.x, w.y, w.z) & ")")
  Call DrawTextHUD(10, 30, 300, 50, "CAM: (" & FmtVal(v.x, v.y, v.z) & ")")
  Call DrawTextHUD(10, 50, 300, 70, "RAE: (" & FmtVal(rr, RtoD * Az, RtoD * el) & ")")
  Call DrawTextHUD(10, 70, 100, 90, "FPS:" & Str$(GetFPS))
End Sub

Private Function FmtVal(ByVal x As Single, ByVal y As Single, ByVal z As Single) As String
  Const FMT As String = " #0.00; -#0.00"
  FmtVal = Format$(x, FMT) & "," & Format$(y, FMT) & "," & Format$(z, FMT)
End Function

Private Sub DrawTitle()
  Call DrawText(200, 10, 640, 50, zTitle)
End Sub

'Draw Text in HUD Font
Private Sub DrawTextHUD(ByVal XLeft As Long, ByVal YTop As Long, _
                        ByVal XRight As Long, ByVal YBottom As Long, _
                        ByRef Text As String, _
                        Optional ByVal ColourRGB As Long = &HFFFFFFFF, _
                        Optional ByVal FormatFlags As Long = 0)
  Dim R As RECT
  R.Left = XLeft
  R.Top = YTop
  R.Right = XRight
  R.bottom = YBottom
  Call zFontHUD.DrawTextW(Text, Len(Text), R, FormatFlags, ColourRGB Or &HFF000000)
End Sub
                    
'---------------------------------------------------------------------------------------------------------
Public Sub Increment()      'A MANUAL VERSION
Attribute Increment.VB_Description = "Move all figures by their deltas, and recalculate positions and Matrices"

 Dim i As Long

  For i = 0 To zNFigures - 1  'Move Figures
    zFigures(i).Increment
  Next i
  zCamera.Increment           'Move Camera
  zViewpoint.Increment        'Move ViewPoint
  For i = 0 To zNLights - 1   'Move Lights
    zLights(i).Increment
  Next i

End Sub

'==========================================================================================================
'================================== REACTION TO MOUSE INPUT ===============================================
'==========================================================================================================

'To Track Mouse movement x- = Left x+ = Right,
'                        y- = Up   y+ = Down,
'                        z- = Out  z+ = In (wheel Towards,Away)   'SORT OF - It doesnt behave intuitively

Public Sub MoveCameraMouse(ByVal ScaledX As Single, ByVal ScaledY As Single, ByVal ScaledZ As Single)
Attribute MoveCameraMouse.VB_Description = "Move camera position by the Mouse"

  Call Mouse.Acquire                                                          'in case it got lost
  Call Mouse.GetDeviceStateMouse(MouseState)                                  'get mouse data
  Call zCamera.MovePosition(MouseState.lX * ScaledX, _
       MouseState.lY * ScaledY, _
       MouseState.lZ * ScaledZ)

End Sub

Public Sub MoveViewpointMouse(ByVal ScaledX As Single, ByVal ScaledY As Single, ByVal ScaledZ As Single)
Attribute MoveViewpointMouse.VB_Description = "Move the viewpoint by mouse data"

  Call Mouse.Acquire                                                          'in case it got lost
  Call Mouse.GetDeviceStateMouse(MouseState)                                  'get mouse data
  Call zViewpoint.MovePosition(MouseState.lX * ScaledX, _
       MouseState.lY * ScaledY, _
       MouseState.lZ * ScaledZ)

End Sub

'To Track Camera movement by Azimuth (y rotation), Elevation (xz rotation), Range (xyz translation)
'       dRange     + Zoom Out , - Zoom In           'metres (Camera Coords)
'       dAzimuth   + Clockwise, - anti clockwise    'In Degrees
'       dElevation + Up       , - Down              'In Degrees

Public Sub MoveCameraRAE(ByVal dRange As Single, ByVal dAzimuth As Single, ByVal dElevation As Single)
Attribute MoveCameraRAE.VB_Description = "Move the camera by change in Range, Azimuth or Elevation"

 Dim Az As Single, el As Single, rr As Single, x As Single, y As Single, z As Single

  With zCamera
    Call RecttoSph(rr, Az, el, .CurrWorldPos.x, .CurrWorldPos.y, .CurrWorldPos.z)          'current position
    el = el + DtoR * dElevation: If el > PI Then el = PI Else If el < -PI Then el = -PI
    Call SphtoRect(x, y, z, rr + dRange, Az - DtoR * dAzimuth, el)   'new position
    Call .SetPosition(x, y, z)
  End With

End Sub

'==========================================================================================================
'================== given a list of keys return the keykode hit if it was in the list, else return -1 =====
'==========================================================================================================
Public Function KeyHit(ParamArray vbKeys() As Variant) As Long
Attribute KeyHit.VB_Description = "Check for a particular Key while scene playing"

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

'-------------------------------- TIMER FUNCTIONS ---------------------------------------------------
Public Sub ResetFPS()                                 'Start FPS Stats
Attribute ResetFPS.VB_Description = "Set Frame counters to zero"

  zNFrames = 0
  Call StopWatchStart(zFrameTimer)

End Sub

Private Sub IncrementFramecount()

  If zNFrames = -1 Then Call ResetFPS
  zNFrames = zNFrames + 1
  If zNFrames > 1000000000 Then Call ResetFPS         'keep it down

End Sub

Public Function GetFPS() As Single                    'Since last ResetFPS
Attribute GetFPS.VB_Description = "Return the current Framerate"
  Dim t As Single
  
  t = StopWatchSplit(zFrameTimer)
  If t > 0 Then GetFPS = Int(10 * zNFrames / t) / 10

End Function

':) Ulli's VB Code Formatter V2.13.5 (01-Feb-03 20:50:58) 90 + 671 = 761 Lines
