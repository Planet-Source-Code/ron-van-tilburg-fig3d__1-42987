VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cLight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

'RVTDX.cLight.cls - a component of the rvtDX.dll
'©2003 Ron van Tilburg - rivit@f1.net.au
'Freeware for Educational Purposes, For commercial interests contact author please, I retain copyright.

'This handles the lighting in a scene

Public Enum LIGHTS_ONOFF
  LIGHT_OFF = 0
  LIGHT_ON = 1
End Enum

Private zIndex  As Long         'My Number
Private zState  As LIGHTS_ONOFF 'My State 1= On, 0=Off
Private zSpecs  As D3DLIGHT8    'My Specs

Private zSpindA As D3DVECTOR    'For Spinning The Light Direction             (Directional and Spot)

Private zRotdA  As D3DVECTOR    'For Moving The Position of Lights (rotational) (point and Spot)
'OR
Private zdPos   As D3DVECTOR    'For Moving The Position of Lights (in x,y,z)   (point and Spot)
'OR BOTH

Private zShineAt As cFigure     'If set follow this Figure  (Directional and Spot)

'======================================= INIT & TERMINATE =================================================
Private Sub Class_Initialize()

  zIndex = -1
  zState = LIGHT_OFF
  Call MakeDirectionalLight

End Sub

Private Sub Class_Terminate()

'


End Sub

'======================== Call These First to Define The Light ============================================
'By Default Shining Down (Sunlight at Noon)
Public Sub MakeDirectionalLight(Optional ByVal Dx As Single = 0, _
                                Optional ByVal Dy As Single = -1, _
                                Optional ByVal Dz As Single = 0)
Attribute MakeDirectionalLight.VB_Description = "Make a light shininh in a particular direction"

  zSpecs.Type = D3DLIGHT_DIRECTIONAL
  Call SetDirection(Dx, Dy, Dz)
  Call SetColourDiffuse
  Call AddMe(zIndex)

End Sub

'By Default Shining Down (Spotlight at y=1)
Public Sub MakeSpotLight(Optional ByVal Px As Single = 0, _
                         Optional ByVal Py As Single = 1, _
                         Optional ByVal Pz As Single = 0, _
                         Optional ByVal Dx As Single = 0, _
                         Optional ByVal Dy As Single = -1, _
                         Optional ByVal Dz As Single = 0, _
                         Optional ByVal A0 As Single = 0, _
                         Optional ByVal A1 As Single = 1, _
                         Optional ByVal A2 As Single = 0, _
                         Optional ByVal Range As Single = 1, _
                         Optional ByVal InnerCone As Single = 30, _
                         Optional ByVal OuterCone As Single = 60, _
                         Optional ByVal Falloff As Single = 1)
Attribute MakeSpotLight.VB_Description = "Make a spotlight"

  zSpecs.Type = D3DLIGHT_SPOT
  Call SetPosition(Px, Py, Pz)
  Call SetDirection(Dx, Dy, Dz)
  Call SetColourDiffuse
  Call SetColourSpecular
  Call SetAttenuation(A0, A1, A2, Range)
  Call SetSpotParms(InnerCone, OuterCone, Falloff)
  Call AddMe(zIndex)

End Sub

'By Default Shining From (LightBulb in all directions at y=1)
Public Sub MakePointLight(Optional ByVal Px As Single = 0, _
                          Optional ByVal Py As Single = 1, _
                          Optional ByVal Pz As Single = 0, _
                          Optional ByVal A0 As Single = 0, _
                          Optional ByVal A1 As Single = 1, _
                          Optional ByVal A2 As Single = 0, _
                          Optional ByVal Range As Single = 1)
Attribute MakePointLight.VB_Description = "Pake a point light like a light bulb"

  zSpecs.Type = D3DLIGHT_POINT
  Call SetPosition(Px, Py, Pz)
  Call SetColourDiffuse
  Call SetColourSpecular
  Call SetAttenuation(A0, A1, A2, Range)
  Call AddMe(zIndex)

End Sub

'======================= POSITION AND DIRECTION =============================================================

'For SPOTLIGHTS & POINTLIGHTS The centre in Scene Coordinates (default is at Y=1)
Public Sub SetPosition(Optional ByVal Px As Single = 0, _
                       Optional ByVal Py As Single = 1, _
                       Optional ByVal Pz As Single = 0)
Attribute SetPosition.VB_Description = "Where a Point or potlight resides"

  Call SetVector(zSpecs.Position, Px, Py, Pz)
  Call AddMe(zIndex)

End Sub

'For SPOTLIGHTS & POINTLIGHTS The centre in Scene Coordinates
Public Sub SetPositionDelta(Optional ByVal Ax As Single = 0, _
                            Optional ByVal Ay As Single = 0, _
                            Optional ByVal Az As Single = 0)
Attribute SetPositionDelta.VB_Description = "How a Point or Spotlight moves"

  Call SetVector(zdPos, Ax, Ay, Az)

End Sub

'The change in Rotational centre in Scene Coordinates (degrees)
Public Sub SetRotationDelta(Optional ByVal dOAx As Single = 0, _
                            Optional ByVal dOAy As Single = 0, _
                            Optional ByVal dOAz As Single = 0)
Attribute SetRotationDelta.VB_Description = "How a Directional or Spotlight rotates (in degrees) per increment"

  Call SetVector(zRotdA, DtoR * dOAx, DtoR * dOAy, DtoR * dOAz)

End Sub

'For All Lights The direction It Shines In From its Scene Position (default is Down)
Public Sub SetDirection(Optional ByVal Dx As Single = 0, _
                        Optional ByVal Dy As Single = -1, _
                        Optional ByVal Dz As Single = 0)
Attribute SetDirection.VB_Description = "In which direction a Directional or Spotlight shines"

  If Dx = 0 And Dy = 0 And Dz = 0 Then Dy = -1
  Call SetVector(zSpecs.Direction, Dx, Dy, Dz)
  Call AddMe(zIndex)

End Sub

'The Rotation of the Figure about its centre of axes (angles supplied in Degrees (0,0,0)=AS IS)
Public Sub SetSpinDelta(Optional ByVal dPAx As Single = 0, _
                        Optional ByVal dPAy As Single = 0, _
                        Optional ByVal dPAz As Single = 0)
Attribute SetSpinDelta.VB_Description = "How a Directional or Spotlight spins (in degrees) per increment"

  Call SetVector(zSpindA, DtoR * dPAx, DtoR * dPAy, DtoR * dPAz)

End Sub

'======================= LIGHT ATTENUATION WITH DISTANCE ===== (SPOTLIGHTS AND POINTLIGHTS) ================

'Model is A = 1/(A0 + A1*D + A2*D*D)  where D is Distance from Light
'Default is Linear Inverse Distance
Private Sub SetAttenuation(Optional ByVal A0 As Single = 0, _
                           Optional ByVal A1 As Single = 1, _
                           Optional ByVal A2 As Single = 0, _
                           Optional ByVal Range As Single = 1)

  With zSpecs
    .Attenuation0 = A0
    .Attenuation1 = A1
    .Attenuation2 = A2
    .Range = Range
  End With
  Call AddMe(zIndex)

End Sub

'======================= SPOTLIGHT CHARACTERISTICS  =========================================================
'Cone size given in Degrees, Linear Falloff between Cones by Default

Private Sub SetSpotParms(Optional ByVal InnerCone As Single = 30, _
                         Optional ByVal OuterCone As Single = 60, _
                         Optional ByVal Falloff As Single = 1)

  OuterCone = Abs(OuterCone)
  InnerCone = Abs(InnerCone)
  If InnerCone = 0 Then InnerCone = 0.1
  If InnerCone > 360 Then InnerCone = 300
  If OuterCone = 0 Then OuterCone = 0.1
  If OuterCone > 360 Then OuterCone = 300
  If OuterCone < InnerCone Then OuterCone = 1.1 * InnerCone
  If Falloff = 0 Then Falloff = 0.01
  With zSpecs
    .Phi = DtoR * OuterCone
    .Theta = DtoR * InnerCone
    .Falloff = Abs(Falloff)
  End With
  Call AddMe(zIndex)

End Sub

'======================= LIGHT COLOURS ===================================================================

'In general components in the Range 0-1 but are allowed to be smaller of greater

'Set a global ambient colour for the Figure (as opposed to a material per vertex) (reaction to Ambient Light)
Public Sub SetColourAmbient(Optional ByVal Red As Single = 1, _
                            Optional ByVal Green As Single = 1, _
                            Optional ByVal Blue As Single = 1)
Attribute SetColourAmbient.VB_Description = "Set the Ambient Light Color"

  With zSpecs.Ambient
    .a = 1
    .R = Red
    .g = Green
    .b = Blue
  End With
  Call AddMe(zIndex)

End Sub

'Set a global diffuse colour for the Figure (as opposed to a material per vertex) (reaction to Lights)
Public Sub SetColourDiffuse(Optional ByVal Red As Single = 1, _
                            Optional ByVal Green As Single = 1, _
                            Optional ByVal Blue As Single = 1)
Attribute SetColourDiffuse.VB_Description = "Set the main colour of the light"

  With zSpecs.Diffuse
    .a = 1
    .R = Red
    .g = Green
    .b = Blue
  End With
  Call AddMe(zIndex)

End Sub

'Set a global specular colour for the Figure (as opposed to a material per vertex) (reaction to Bright Lights)
Public Sub SetColourSpecular(Optional ByVal Red As Single = 1, _
                             Optional ByVal Green As Single = 1, _
                             Optional ByVal Blue As Single = 1)
Attribute SetColourSpecular.VB_Description = "Set the bright colour component"

  With zSpecs.specular
    .a = 1
    .R = Red
    .g = Green
    .b = Blue
  End With
  Call AddMe(zIndex)

End Sub

'returns new state
Public Property Let State(ByVal NewState As LIGHTS_ONOFF)
Attribute State.VB_Description = "Are we ON or OFF"

  zState = NewState
  On Local Error Resume Next
  If zIndex <> -1 Then Call D3DD.LightEnable(zIndex, zState)
  If Err.Number <> D3D_OK Then                                  'Code to handle the error goes here.
  End If
  On Error GoTo 0

End Property

Public Property Get State() As LIGHTS_ONOFF

  State = zState

End Property

Public Property Set ShineAt(ByRef Target As cFigure)
Attribute ShineAt.VB_Description = "Aim a spotlight or Directional light at this Figure - it will track it"

  Set zShineAt = ShineAt

End Property

Public Property Get ShineAt() As cFigure

  Set ShineAt = zShineAt

End Property

Public Sub Increment()                                                              'For Moving Lights around
Attribute Increment.VB_Description = "How should this light move per increment"

 Dim tmpMat As D3DMATRIX, tmpVec As D3DVECTOR, Update As Boolean

  Select Case zSpecs.Type
  Case D3DLIGHT_POINT:       'Can be moved around but cant point
    Call IncrementVector(zSpecs.Position, zdPos)                                    'Positional Move
    tmpVec = zSpecs.Position
    Call D3DXMatrixRotationYawPitchRoll(tmpMat, zRotdA.y, zRotdA.x, zRotdA.z)       'Rotated about (0,0,0)
    Call D3DXVec3TransformCoord(zSpecs.Position, tmpVec, tmpMat)
    Update = True

  Case D3DLIGHT_SPOT:        'Can be moved around and Point
    If zShineAt Is Nothing Then
      tmpVec = zSpecs.Direction
      Call D3DXMatrixRotationYawPitchRoll(tmpMat, zSpindA.y, zSpindA.x, zSpindA.z)  'Rotated about YXZ
      Call D3DXVec3TransformCoord(zSpecs.Direction, tmpVec, tmpMat)
      Update = True
    Else
      zSpecs.Direction = VectorDiff(zShineAt.CurrWorldPos, zSpecs.Position)        'Set the Vector Equal to the Position
      Update = True
    End If

  Case D3DLIGHT_DIRECTIONAL: 'Cant be moved around but can point
    If zShineAt Is Nothing Then
      tmpVec = zSpecs.Direction
      Call D3DXMatrixRotationYawPitchRoll(tmpMat, zSpindA.y, zSpindA.x, zSpindA.z)  'Rotated about YXZ
      Call D3DXVec3TransformCoord(zSpecs.Direction, tmpVec, tmpMat)
      Update = True
    Else
      zSpecs.Direction = zShineAt.CurrWorldPos                                      'Set the Vector Equal to the Position
      Update = True
    End If
  End Select

  If Update Then Call AddMe(zIndex)

End Sub

Friend Sub AddMe(ByVal Index As Long)

  If Index <> -1 Then
    zIndex = Index
    Call D3DD.SetLight(Index, zSpecs)
  End If

End Sub

':) Ulli's VB Code Formatter V2.13.5 (31-Jan-03 12:21:13) 19 + 294 = 313 Lines
