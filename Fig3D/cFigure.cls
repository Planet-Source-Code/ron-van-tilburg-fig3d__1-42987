VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFigure"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'RVTDX.cFigure - a component of the rvtDX.dll
'©2003 Ron van Tilburg - rivit@f1.net.au
'Freeware for Educational Purposes, For commercial interests contact author please, I retain copyright.

'This handles the Figure components of a Scene

Public Enum FIG_MAT
  FIGM_VXC = 0        'Colours on Vertices
  FIGM_MAT = 1        'Use Material
End Enum

Public Type TexSpec
  Index       As Integer
  RenderMode  As Integer
End Type

Private zFigSpec      As FIG_SPEC
Private zRenderFlags  As RENDER_FLAGS

Private zParent       As cFigure     'This Figure's Parent (well behaved children move about their Parents)

Private zCurrWorldPos As D3DVECTOR   'Where we live Right Now
Private zPrevWorldPos As D3DVECTOR   'Where we were before the Last Increment

'Initial Axial Parameters
Private zAxiRotA      As D3DVECTOR   'Spin about its own Centre            (PAx,PAy,PAz) default = (0,0,0)
Private zAxiRotdA     As D3DVECTOR   'Increment angles about axes       (dPAx,dPAy,dPAz) default = (0,0,0)

Private zAxiScale     As D3DVECTOR   'Scales in x,y,z directions           (PSx,PSy,PSz) default = (1,1,1)
Private zAxidScale    As D3DVECTOR   'Increment in scale                (dPSx,dPSy,dPSz) default = (0,0,0)

Private zAxiTiltA     As D3DVECTOR   'Tilt Of Figure                       (Pix,Piy,Piz) default = (0,0,0)

'Initial Orbital Parameters
Private zOrbOrigin    As D3DVECTOR   'Scene Rotational Centre pos            (Ox,Oy,Oz)  default = (0,0,0)
Private zOrbdOrigin   As D3DVECTOR   'Change on Centre of Rotation         (dOx,dOy,dOz) default = (0,0,0)

Private zOrbPos       As D3DVECTOR   'Scene Centre of Figure                 (Px,Py,Pz)  default = (0,0,0)
Private zOrbdPos      As D3DVECTOR   'Scene Centre of Figure               (dPx,dPy,dPz) default = (0,0,0)

Private zOrbRotA      As D3DVECTOR   'Scene Rotational Angle               (OAx,OAy,OAz) default = (0,0,0)
Private zOrbRotdA     As D3DVECTOR   'Scene Rotational Angle Delta      (dOAx,dOAy,dOAz) default = (0,0,0)

Private zOrbTiltA     As D3DVECTOR   'Tilt Of Orbit                        (Oix,Oiy,Oiz) default = (0,0,0)

Private zOrbMatrix    As D3DMATRIX   'The Most Current Cumulative Orbital Transformtion Matrix
Private zOrbMatrixOK  As Boolean     'TRUE if we can rely on the Transformation Matrix

'Colours, Material and Textures
Private zMatMode    As FIG_MAT
Private zMaterial   As D3DMATERIAL8  'Color information for the entire figure as one

Private zTextures(0 To 7) As TexSpec 'These are indices into Scene.Texture(i)
'can only store up to DD3DC.MaxSimultaneousTextures to be mapped onto our figure
Private zNTextures        As Integer 'The Number We will Use

Private zFigSpecIndex     As Long    'The Index to Scene.FigSpecs(i)

'Items for FIGS_TEXT Figures
Private zText         As String
Private zTextFont     As StdFont
Private zMesh         As D3DXMesh    'IF a TEXT object, Or a TEAPOT else NOTHING

'===========================================================================================================
Private Sub Class_Initialize()

 Dim i As Long

  zFigSpec = FIGS_CUBE
  Call SetPosition
  Call SetScale(1, 1, 1)
  
  For i = 0 To 7          'Mark Textures as Unused
    zTextures(i).Index = -1
  Next i

End Sub

Private Sub Class_Terminate()

  Set zParent = Nothing
  Set zTextFont = Nothing
  Set zMesh = Nothing
End Sub

'===================== PROPERTIES ==========================================================================

Public Property Let FigSpec(ByVal FigSpec As FIG_SPEC)
Attribute FigSpec.VB_Description = "The Specification of a given Figure (see cDX8.cls for more details)"

 Dim FT As Long

  If (FigSpec And FIGS_CUSTOM) = 0 Then
    FT = FigSpec And FIGS_TYPEMASK
    If FT >= FIGS_LOCATOR And FT <= FIGS_LAST Then
      zFigSpec = FigSpec
      If FT = zFigSpec Then                                               'Absolutely no parameters given
        P1 = 0: P2 = 0: P3 = 0                                            'Fill some Defaults
      End If
    Else
      zFigSpec = FIGS_CUBE
    End If
  Else
    zFigSpec = FigSpec                'Dont check anything for CUSTOM Figures
  End If

End Property

Public Property Let P1(ByVal P1 As Single)                              'Parameter 1 usually N
Attribute P1.VB_Description = "Parameter 1 for Figure Generation"

  If (zFigSpec And FIGS_CUSTOM) = 0 Then                                  'only parse for stock objects
    Select Case zFigSpec And FIGS_TYPEMASK
    Case FIGS_LOCATOR, FIGS_POINT, FIGS_LINE:
      P1 = 0
    Case FIGS_AXES, FIGS_AXISPLANES, FIGS_MESH:
      Exit Property
    Case FIGS_SHEET:
      P1 = Abs(P1)
    Case Else:                                                            'N
      P1 = Int(P1)
      If P1 < 2 Then P1 = 2
      If P1 > 255 Then P1 = 255
    End Select
  End If
  zFigSpec = (zFigSpec And Not FIGS_P1MASK) Or (CLng(P1) And &HFF) * 256&

End Property

Public Property Let P2(ByVal P2 As Single)                          'Parameter 1 usually M,Angle or ratio
Attribute P2.VB_Description = "Parameter 2 for Figure Generation"

  If (zFigSpec And FIGS_CUSTOM) = 0 Then                                  'only parse for stock objects
    Select Case zFigSpec And FIGS_TYPEMASK
    Case FIGS_POLYARC, FIGS_POLYWEDGE:                                'Angle 0-358.59
      If P2 < 0 Then P2 = 0
      P2 = Fix(256 * P2 / 360)
      If P2 > 255 Then P2 = 255
    Case FIGS_POLYMESH, FIGS_SPHEROID, FIGS_TOROID, FIGS_ASTROID:        'M
      P2 = Int(P2)
      If P2 < 2 Then P2 = 2
      If P2 > 255 Then P2 = 255
    Case FIGS_POLYSTAR, FIGS_POLYWASHER, FIGS_FRUSTRUM, FIGS_PIPE:        'A Ratio 0-0.99609
      If P2 < 0 Then P2 = 0
      If P2 > 1 Then P2 = 255 Else P2 = Int(256 * P2)
    Case FIGS_SHEET:
      P2 = Abs(P2)
    Case Else:
      P2 = 0
    End Select
    zFigSpec = (zFigSpec And Not FIGS_P2MASK) Or (CLng(P2) And &HFF) * 65536
  End If

End Property

Public Property Let P3(ByVal P3 As Single)                          'Parameter 3 usually Ratio
Attribute P3.VB_Description = "Parameter three for Figure Generation"

  If (zFigSpec And FIGS_CUSTOM) = 0 Then                                  'only parse for stock objects
    Select Case zFigSpec And FIGS_TYPEMASK
    Case FIGS_TOROID, FIGS_SHEET:                                         'A Ratio 0-0.99609
      If P3 < 0 Then P3 = 0
      If P3 > 1 Then P3 = 255 Else P3 = Int(256 * P3)
    Case Else:
      P3 = 0
    End Select
  End If
  If CLng(Abs(P3) < 128) Then
    zFigSpec = (zFigSpec And Not FIGS_P3MASK) Or (CLng(Abs(P3)) And &HFF) * 16777216
  Else
    zFigSpec = (zFigSpec And Not FIGS_P3MASK) Or (CLng(Abs(P3)) - 128 And &HFF) * 16777216 Or &H80000000
  End If

End Property

'Only used for FIGS_TEXT object types
Public Property Let Text(ByRef NewText As String)
Attribute Text.VB_Description = "For Figures made of Letters FIGS_TEXT,  the text of those letters"
  zText = NewText
End Property

Public Property Set TextFont(ByRef NewFont As StdFont)
Attribute TextFont.VB_Description = "Font to Use when generating a FIGS_TEXT figure (only)"
  Set zTextFont = NewFont
End Property

Friend Function GenerateMesh() As Boolean
  
  If (zFigSpec And FIGS_TYPEMASK) = FIGS_MESH Then        'Generate a Mesh
    If zFigSpec = FIGS_TEXT Then                          'Generate a Mesh
      Set zMesh = CreateTextMesh(zTextFont, zText)        'In mDXEngine
      GenerateMesh = Not (zMesh Is Nothing)
    ElseIf zFigSpec = FIGS_TEAPOT Then                    'Generate a Mesh
      Set zMesh = D3DX.CreateTeapot(D3DD, Nothing)
      GenerateMesh = Not (zMesh Is Nothing)
    End If
  Else
    'any other Types fail
  End If
End Function

Friend Property Get Mesh() As D3DXMesh
  Set Mesh = zMesh
End Property
'-----------------------------------------------------------------------------------------------------------
Public Property Get FigSpec() As FIG_SPEC

  FigSpec = zFigSpec

End Property

Public Sub AttachTexture(ByVal PipelineStage As Long, _
                         ByVal SceneTextureIndex As Long, _
                         Optional ByVal RenderMode As RENDER_FLAGS = 0)
Attribute AttachTexture.VB_Description = "Associate a Texture (Index) with this texture"

  If zNTextures < D3DDC.MaxSimultaneousTextures Then 'The Device can only support this many attached to blending operations
    If PipelineStage >= 0 And PipelineStage <= 7 Then
      zTextures(PipelineStage).Index = SceneTextureIndex
      zTextures(PipelineStage).RenderMode = RenderMode
    End If
  End If

End Sub

Public Property Get TextureIndex(ByVal PipelineStage As Long) As Long
Attribute TextureIndex.VB_Description = "Texture nr associated with a Scene Texture and used by this Figure ."

  If PipelineStage >= 0 And PipelineStage <= 7 Then
    TextureIndex = zTextures(PipelineStage).Index
  End If

End Property

Public Property Get TextureRenderMode(ByVal PipelineStage As Long) As Long
Attribute TextureRenderMode.VB_Description = "How a prticular Texture should be additionally treated, ovverrides Figure and Scene RenderFlags"

  If PipelineStage >= 0 And PipelineStage <= 7 Then
    TextureRenderMode = zTextures(PipelineStage).RenderMode
  End If

End Property

Friend Property Get Material() As D3DMATERIAL8

  Material = zMaterial

End Property

Friend Property Let FigSpecIndex(ByRef FigSpecIndex As Long)

  zFigSpecIndex = FigSpecIndex

End Property

Friend Property Get FigSpecIndex() As Long

  FigSpecIndex = zFigSpecIndex

End Property

Friend Property Get MatMode() As FIG_MAT

  MatMode = zMatMode

End Property

Public Property Set Parent(ByRef NewParent As cFigure)
Attribute Parent.VB_Description = "A figure for this figure to follow. Give the Child Deltas relative to this Parent"

  Set zParent = NewParent

End Property

Public Property Get Parent() As cFigure

  Set Parent = zParent

End Property

Public Property Let RenderFlags(ByVal NewFlags As RENDER_FLAGS)    'Will Ovverride Scene RenderFlags
Attribute RenderFlags.VB_Description = "Render Parameters for a figure - these override scene parameters"
  zRenderFlags = NewFlags
End Property

Public Property Get RenderFlags() As RENDER_FLAGS
  RenderFlags = zRenderFlags
End Property

'===================== METHODS =============================================================================

'===================== MATERIALS   (Figure Colours) ========================================================
'In general components in the Range 0-1 but are allowed to be smaller or greater

'Set a global ambient colour for the Figure (as opposed to a material per vertex) (reaction to Ambient Light)
Public Sub SetColourAmbient(Optional ByVal Alpha As Single = 1, _
                            Optional ByVal Red As Single = 1, _
                            Optional ByVal Green As Single = 1, _
                            Optional ByVal Blue As Single = 1)
Attribute SetColourAmbient.VB_Description = "Set the colour for reacting woth ambient color"

  zMatMode = FIGM_MAT
  With zMaterial.Ambient
    .a = Alpha
    .R = Red
    .g = Green
    .b = Blue
  End With

End Sub

'Set a global diffuse colour for the Figure (as opposed to a material per vertex) (reaction to Lights)
Public Sub SetColourDiffuse(Optional ByVal Alpha As Single = 1, _
                            Optional ByVal Red As Single = 1, _
                            Optional ByVal Green As Single = 1, _
                            Optional ByVal Blue As Single = 1)
Attribute SetColourDiffuse.VB_Description = "Set the color for reacting with light ingeneral"

  zMatMode = FIGM_MAT
  With zMaterial.Diffuse
    .a = Alpha
    .R = Red
    .g = Green
    .b = Blue
  End With

End Sub

'Set a global specular colour for the Figure (as opposed to a material per vertex) (reaction to Bright Lights)
Public Sub SetColourSpecular(Optional ByVal Alpha As Single = 1, _
                             Optional ByVal Red As Single = 1, _
                             Optional ByVal Green As Single = 1, _
                             Optional ByVal Blue As Single = 1, _
                             Optional ByVal Power As Single = 1)
Attribute SetColourSpecular.VB_Description = "Set the colour under influence of bright lights"

  zMatMode = FIGM_MAT
  With zMaterial.specular
    .a = Alpha
    .R = Red
    .g = Green
    .b = Blue
  End With
  zMaterial.Power = Power   'how strong the effect is (set to Zero to Turn it off)

End Sub

'Set a global colour for the Figure (as opposed to a material per vertex) (glow reaction of itself)
Public Sub SetColourEmissive(Optional ByVal Alpha As Single = 1, _
                             Optional ByVal Red As Single = 1, _
                             Optional ByVal Green As Single = 1, _
                             Optional ByVal Blue As Single = 1)
Attribute SetColourEmissive.VB_Description = "Set the colour for apparent glow"

  zMatMode = FIGM_MAT
  With zMaterial.emissive
    .a = Alpha
    .R = Red
    .g = Green
    .b = Blue
  End With

End Sub

'===================================== Positioning of Figures ===============================================
'The centre in Scene Coordinates
Public Sub SetPosition(Optional ByVal Px As Single = 0, _
                       Optional ByVal Py As Single = 0, _
                       Optional ByVal Pz As Single = 0)
Attribute SetPosition.VB_Description = "Set the initial XYZ for a figure"

  Call SetVector(zOrbPos, Px, Py, Pz)
  zOrbMatrixOK = False

End Sub

'The centre in Scene Coordinates
Public Sub SetPositionRAE(ByVal Range As Single, ByVal Azimuth As Single, ByVal Elevation As Single)
Attribute SetPositionRAE.VB_Description = "Set the position by Range, Azimuth (clockwise) and Elevation (from XZ)"

  Call SphtoRect(zOrbPos.x, zOrbPos.y, zOrbPos.z, Range, Azimuth, Elevation)
  zOrbMatrixOK = False

End Sub

Public Property Get CurrWorldPos() As D3DVECTOR   'Fully treated for all components
Attribute CurrWorldPos.VB_Description = "Return the Position right Now in World Coordinates"

  If Not zOrbMatrixOK Then
    ReCalcOrbMatrix
    ReCalcCurrPosition
  End If
  CurrWorldPos = zCurrWorldPos

End Property

Public Property Get PrevWorldPos() As D3DVECTOR
Attribute PrevWorldPos.VB_Description = "Where the figure was last increment"

  PrevWorldPos = zPrevWorldPos

End Property

Public Property Get InitialPosition() As D3DVECTOR         'This is always the Initial Position
Attribute InitialPosition.VB_Description = "Return the initial XYZ coordinates for a figure"

  InitialPosition = zOrbPos

End Property

Friend Property Get OrbMatrix() As D3DMATRIX

  If Not zOrbMatrixOK Then
    ReCalcOrbMatrix
    ReCalcCurrPosition
  End If
  OrbMatrix = zOrbMatrix

End Property

Public Sub MovePosition(ByVal dPx As Single, ByVal dPy As Single, ByVal dPz As Single)
Attribute MovePosition.VB_Description = "Move Figure in World by d XYZ"

 Dim tmpV As D3DVECTOR

  Call SetVector(tmpV, dPx, dPy, dPz)
  If Not IsZeroVector(tmpV) Then                           'only if it moved
    Call IncrementVector(zOrbPos, tmpV)
    zOrbMatrixOK = False
  End If

End Sub

'The change of centre in Scene Coordinates
Public Sub SetPositionDelta(Optional ByVal dPx As Single = 0, _
                            Optional ByVal dPy As Single = 0, _
                            Optional ByVal dPz As Single = 0)
Attribute SetPositionDelta.VB_Description = "Set by how much the XYZ position will move per increment"

  Call SetVector(zOrbdPos, dPx, dPy, dPz)

End Sub

'The Rotational centre in Scene Coordinates
Public Sub SetOrbitOrigin(Optional ByVal Ox As Single = 0, _
                          Optional ByVal Oy As Single = 0, _
                          Optional ByVal Oz As Single = 0)
Attribute SetOrbitOrigin.VB_Description = "Set the centre of a figures orbit - by default (0,0,0)"

  Call SetVector(zOrbOrigin, Ox, Oy, Oz)
  zOrbMatrixOK = False

End Sub

'The change in Rotational centre in Scene Coordinates
Public Sub SetOrbitOriginDelta(Optional ByVal dOx As Single = 0, _
                               Optional ByVal dOy As Single = 0, _
                               Optional ByVal dOz As Single = 0)
Attribute SetOrbitOriginDelta.VB_Description = "The amount to move the Orbit Origin per frame"

  Call SetVector(zOrbdOrigin, dOx, dOy, dOz)

End Sub

'The Rotational centre in Scene Coordinates
Public Sub SetRotation(Optional ByVal OAx As Single = 0, _
                       Optional ByVal OAy As Single = 0, _
                       Optional ByVal OAz As Single = 0)
Attribute SetRotation.VB_Description = "Set how far in its orbit(degrees)  the figure should start"

  Call SetVector(zOrbRotA, DtoR * OAx, DtoR * OAy, DtoR * OAz)
  zOrbMatrixOK = False

End Sub

'The change in Rotational centre in Scene Coordinates
Public Sub SetRotationDelta(Optional ByVal dOAx As Single = 0, _
                            Optional ByVal dOAy As Single = 0, _
                            Optional ByVal dOAz As Single = 0)
Attribute SetRotationDelta.VB_Description = "How fast a figure should rotate (degrees)  in its orbit per increment"

  Call SetVector(zOrbRotdA, DtoR * dOAx, DtoR * dOAy, DtoR * dOAz)

End Sub

'The Tilt of The Orbit
Public Sub SetOrbitalTilt(Optional ByVal Oix As Single = 0, _
                          Optional ByVal Oiy As Single = 0, _
                          Optional ByVal Oiz As Single = 0)
Attribute SetOrbitalTilt.VB_Description = "Set the angle at which the Figure orbit inclines from XZ plane"

  Call SetVector(zOrbTiltA, DtoR * Oix, DtoR * Oiy, DtoR * Oiz)
  zOrbMatrixOK = False

End Sub

'The Rotation of the Figure about its centre (angles supplied in Degrees (0,0,0)=AS IS)
Public Sub SetSpin(Optional ByVal PAx As Single = 0, _
                   Optional ByVal PAy As Single = 0, _
                   Optional ByVal PAz As Single = 0)
Attribute SetSpin.VB_Description = "Set the amount by which the figure has spun (degrees)  on its axis"

  Call SetVector(zAxiRotA, DtoR * PAx, DtoR * PAy, DtoR * PAz)
  zOrbMatrixOK = False

End Sub

'The Rotation of the Figure about its centre (angles supplied in Degrees (0,0,0)=AS IS)
Public Sub SetSpinDelta(Optional ByVal dPAx As Single = 0, _
                        Optional ByVal dPAy As Single = 0, _
                        Optional ByVal dPAz As Single = 0)
Attribute SetSpinDelta.VB_Description = "How fast a figure should spin (in degrees) per increment"

  Call SetVector(zAxiRotdA, DtoR * dPAx, DtoR * dPAy, DtoR * dPAz)

End Sub

'The Tilt of The Figure
Public Sub SetTilt(Optional ByVal Pix As Single = 0, _
                   Optional ByVal Piy As Single = 0, _
                   Optional ByVal Piz As Single = 0)
Attribute SetTilt.VB_Description = "How far is the Figure Tilted (degrees) from its XZ plane"

  Call SetVector(zAxiTiltA, DtoR * Pix, DtoR * Piy, DtoR * Piz)
  zOrbMatrixOK = False

End Sub

'The Figure's Scale
Public Sub SetScale(Optional ByVal PSx As Single = 1, _
                    Optional ByVal PSy As Single = 1, _
                    Optional ByVal PSz As Single = 1)
Attribute SetScale.VB_Description = "How large a figure should be.Std Cubes are side= 1,Std  Spheres Radius= 1"

  Call SetVector(zAxiScale, PSx, PSy, PSz)

End Sub

'The change of scale
Public Sub SetScaleDelta(Optional ByVal dPSx As Single = 0, _
                         Optional ByVal dPSy As Single = 0, _
                         Optional ByVal dPSz As Single = 0)
Attribute SetScaleDelta.VB_Description = "Set the amount by which size should change per increment"

  Call SetVector(zAxidScale, dPSx, dPSy, dPSz)

End Sub

'Prepare a set of vertices based on Figure Parameters, prepared to correct position and scale in World Space
'When called returns the World Matrix for This Figure
'If AutoIncrement is True then we update the various parameters of motion 1 step (for next time)

'                               +     Scale (+ScaleDelta)
'                               +++   Spin  (+SpinDelta)
'                                 +   Tilt
'                              /
'                             /
'                            /
'                           /                                         Figures  Move by dPosition
'                          /                                                    or  by dRotation
'                         /
'                        / Rotation
'                       / (+RotationDelta)
'                      /  OrbitTilt                   OR
'       OrbitOrigin   O ---------------------- +  Position (+PositionDelta)
'
Friend Function GetWorldMatrix() As D3DMATRIX

'Modify the Figure  ie. Scale, Rotate, Tilt

  Call MakeFigureMatrix(GetWorldMatrix, zAxiScale, zAxiRotA, zAxiTiltA)           'in mDXEngine

  'Place it into scene Position  ie.'Translate, Rotate, Tilt, Translate Orbit
  Call ReCalcOrbMatrix

  'Turn it into The World Matrix
  Call D3DXMatrixMultiply(GetWorldMatrix, GetWorldMatrix, zOrbMatrix)             'The Final World Matrix

  zPrevWorldPos = zCurrWorldPos                                                   'We were Here
  Call ReCalcCurrPosition                                                         'We are Now Here

End Function

'NOTE::: THAT INCREMENTING DOES NOT INVALIDATE ORBMATRIX
'IT IS ASSUMED INCREMENTING IS THE LAST STEP TAKEN IN A SCENE RENDER LOOP
'AND GETWORLDMATRIX() WILL REDO IT ANYWAY. Not Only that RENDERING AND MATHS ARE NOW IN SYNCH

Public Sub Increment()
Attribute Increment.VB_Description = "Make every moving figure move by One Delta"

  Call IncrementVector(zOrbPos, zOrbdPos)                     'Fix new Figure Position
  Call IncrementVector(zOrbRotA, zOrbRotdA)                   'Fix new Orbital Rotation about its origin
  Call IncrementVector(zOrbOrigin, zOrbdOrigin)               'Fix new Orbit about its origin
  Call IncrementVector(zAxiRotA, zAxiRotdA)                   'Fix new Figure Spin
  Call IncrementVector(zAxiScale, zAxidScale)                 'Fix new Figure Scale

  'However as Camera and ViewPoint Figures never actually Get Rendered we update their zOrbMatrix now
  If (zFigSpec And FIGS_TYPEMASK) = FIGS_LOCATOR Then
    Call ReCalcOrbMatrix
    zPrevWorldPos = zCurrWorldPos                                                   'We were Here
    Call ReCalcCurrPosition                                                         'We are Now Here
  End If

End Sub

Private Sub ReCalcOrbMatrix()

'Place it into scene Position  ie.'Translate, Rotate, Tilt, Translate Orbit

  Call MakeOrbitMatrix(zOrbMatrix, zOrbPos, zOrbRotA, zOrbTiltA, zOrbOrigin)      'in mDXEngine

  'If there is a Parent Do the Final Translation To its Position (which is now correct we hope)
  If Not zParent Is Nothing Then
    Call D3DXMatrixMultiply(zOrbMatrix, zOrbMatrix, zParent.OrbMatrix)            'Update Our Orbit Matrix
  End If

  zOrbMatrixOK = True

End Sub

Private Sub ReCalcCurrPosition()

  Call D3DXVec3TransformCoord(zCurrWorldPos, WORLD_ORIGIN, zOrbMatrix)                'We are Now Here

End Sub

'===========================================================================================================

':) Ulli's VB Code Formatter V2.13.5 (01-Feb-03 20:52:14) 59 + 504 = 563 Lines
