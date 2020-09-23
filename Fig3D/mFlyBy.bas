Attribute VB_Name = "mFlyBy"
Option Explicit

'Fig3D - a Demonstration of the capabiltites of the rvtDX.dll  D3D graphics engine
'©2003 Ron van Tilburg - rivit@f1.net.au
'Freeware for Educational Purposes, For commercial interests contact author please, I retain copyright.

'=========================== A DEMO OF A FLYBY OF SATURN - You are Voyager ================================

'With a bit of luck you fly between the planet and the Rings
'the camera shows where you are going and then where you have been
'It will run for about 4500 frames or 90secs @ 50fps
'TimeScale is: 2400 Frames is 1 Day=24 hrs ie. 1hr=100 Frames
'In this Time Voyager is doing a nominal (my guess) 5.65km/sec or 1 Frame = 2036 KM

'Saturn is about 1425 Million km from Earth, Hitting this was a brilliant piece of engineering
'As it progresses you will see just how hard it must have been....
'BTW Lighting has been upped to be able to see anything - in real life its pretty black out there

Public Sub ShowSaturnFlyby()

 Const SA_EQ As Single = 26.75
 Const SA_SPIN As Single = -360 / 1065
 Const CAM_DX As Single = -0.101575
 Const CAM_DY As Single = -0.175933

 Dim Scene      As cScene            'Where we will Place our Figures
 Dim Figure     As cFigure           'a useful helper
 Dim Light      As cLight            'another useful helper

  Set Scene = New cScene             'Create a scene (always good fun)
  With Scene                         'Now fill it in
    .Title = "Saturn Flyby"
    .ShowHUD = True
    Call .SetAmbientLight(32, 32, 32)               'Ambient Light is pretty Dark (Its along way from the Sun)

    'We need a distant LightSource (SUN)                  'Add a directional light that is very bright
    Set Light = New cLight
    With Light
      Call .MakeDirectionalLight(1, 0, -1.15)
      Call .SetColourDiffuse(6, 6, 6)
    End With
    Call .AddLight(Light)                                 'This turns it ON

    'Setup Camera                                         'Camera is Voyager
    With .Camera
      Call .SetPosition(297, 0, 527.8)                    '304.75
      Call .SetPositionDelta(CAM_DX, 0, CAM_DY)           'we Hit it about Frame 3000
    End With

    'Now make figures                                             'We need only three Textures
    Call .TextureFromFile(App.Path & "\textures\Saturn.bmp")           '0
    Call .TextureFromFile(App.Path & "\textures\SaturnRings.bmp")      '1
    Call .TextureFromFile(App.Path & "\textures\SaturnTitan.bmp")      '2

    'Two StarFields                                       'An example of Billboarding illusion
    Set Figure = New cFigure                              'First Field Behind where we start
    With Figure
      .FigSpec = FIGS_POINTFIELD
      .P1 = 5                                             '500 stars
      Call .SetScale(700, 0, 700)                         'Make it pretty Large
      Call .SetPosition(410, 0, 700)                      'far enough behind not to hit Saturn
      Call .SetPositionDelta(CAM_DX, 0, CAM_DY)           'move it up behind us
      Call .SetTilt(90, 60, 0)                            'Vertical
      Call .AttachTexture(0, 0)
    End With
    Call .AddFigure(Figure)

    Set Figure = New cFigure                              'Second Field Behind Saturn
    With Figure
      .FigSpec = FIGS_POINTFIELD
      .P1 = 10                                            '500 stars
      Call .SetScale(3500, 0, 3000)                       'Make it pretty Large
      Call .SetPosition(-110, 0, -170)
      Call .SetPositionDelta(CAM_DX, 0, CAM_DY)           'move it away as we approach
      Call .SetTilt(90, 60, 0)                            'Vertical
      Call .AttachTexture(0, 0)
    End With
    Call .AddFigure(Figure)

    'Tethys                                               'Top right When we leave
    Set Figure = New cFigure
    With Figure
      .FigSpec = FIGS_SPHERE
      .P1 = 16
      .P2 = 8
      Call .SetScale(0.0525, 0.0525, 0.0525)              '525KM Radius
      Call .SetPosition(29.5, 0, 0)                       'Position it Here
      Call .SetOrbitalTilt(SA_EQ + 1.1, 0, 0)             'Inclined from Centre
      Call .SetSpinDelta(0, SA_SPIN, 0)                   'Spin about y axis by these degrees each frame
      Call .SetRotation(0, 315, 0)                        '169 Abitrary
      Call .SetRotationDelta(0, -360 / 4531.2, 0)         '1.888d = 2400*1.888 = 4531.2 Frames
      Call .AttachTexture(0, 0)
    End With
    Call .AddFigure(Figure)

    'Dione                                                '2nd from Bottom on approach
    Set Figure = New cFigure
    With Figure
      .FigSpec = FIGS_SPHERE
      .P1 = 16
      .P2 = 8
      Call .SetScale(0.056, 0.056, 0.056)                 '560KM
      Call .SetPosition(37.7, 0, 0)                       'Position it Here
      Call .SetOrbitalTilt(SA_EQ, 0, 0)                   'Inclined from Centre
      Call .SetSpinDelta(0, SA_SPIN, 0)                   'Spin about y axis by these degrees each frame
      Call .SetRotation(0, 72, 0)                         '107 Abitrary
      Call .SetRotationDelta(0, -360 / 6568.8, 0)         '2.737d = 2400*2.737 = 6568.8 Frames
      Call .AttachTexture(0, 0)
    End With
    Call .AddFigure(Figure)

    'Rhea                                                 'at bottom on approach
    Set Figure = New cFigure
    With Figure
      .FigSpec = FIGS_SPHERE
      .P1 = 16
      .P2 = 8
      Call .SetScale(0.0765, 0.0765, 0.0765)              '765KM
      Call .SetPosition(52.7, 0, 0)                       'Position it Here
      Call .SetOrbitalTilt(SA_EQ + 0.4, 0, 0)             'Inclined from Centre
      Call .SetSpinDelta(0, SA_SPIN, 0)                   'Spin about y axis by these degrees each frame
      Call .SetRotation(0, 340, 0)                        '70 Abitrary
      Call .SetRotationDelta(0, -360 / 10843.2, 0)        '4.518d = 2400*4.518 = 10843.2 Frames
      Call .AttachTexture(0, 0)
    End With
    Call .AddFigure(Figure)

    'Japetus                                              'havent found this one yet...
    Set Figure = New cFigure
    With Figure
      .FigSpec = FIGS_SPHERE
      .P1 = 16
      .P2 = 8
      Call .SetScale(0.072, 0.072, 0.072)                 '720KM
      Call .SetPosition(356, 0, 0)                        'Position it Here
      Call .SetOrbitalTilt(SA_EQ + 14.7, 0, 0)            'Inclined from Centre
      Call .SetSpinDelta(0, SA_SPIN, 0)                   'Spin about y axis by these degrees each frame
      Call .SetRotation(0, 66.6, 0)                       '156 Abitrary
      Call .SetRotationDelta(0, -360 / 190392, 0)         '79.33d = 2400*79.33 = 190392 Frames
      Call .AttachTexture(0, 0)
    End With
    Call .AddFigure(Figure)

    'Saturn Body
    Set Figure = New cFigure
    With Figure
      .FigSpec = FIGS_SPHERE
      Call .SetScale(6, 5.4, 6)                           '60,000KM Radius, 1/10 flattening
      Call .SetTilt(SA_EQ, 0, 0)                          '26,44 Inclined from Centre
      Call .SetSpinDelta(0, SA_SPIN, 0)                   '10h 39m => 360/(100*10.65)
      Call .SetColourDiffuse(0.8, 1, 1, 1)
      Call .AttachTexture(0, 0)
    End With
    Call .AddFigure(Figure)

    'Saturn Rings
    Set Figure = New cFigure
    With Figure
      .FigSpec = FIGS_WASHER
      .P1 = 72
      .P2 = 67 / 140.6                                    'Inner Radius/Outer Radius
      Call .SetScale(14.06, 14.06, 14.06)                 '67,000KM - 140,600KM Radius C->F Ring
      Call .SetTilt(SA_EQ, 0, 0)                          'Inclined from Centre
      Call .SetSpinDelta(0, SA_SPIN, 0)                   '10h 39m => 360/(100*10.65)
      Call .AttachTexture(0, 1)                           'apply this twice to get the required transparency
      Call .AttachTexture(1, 1)
      .RenderFlags = RF_LIGHTTINT
    End With
    Call .AddFigure(Figure)

    'Titan                                                'Far right on approach, Mid Left on departure
    Set Figure = New cFigure
    With Figure
      .FigSpec = FIGS_SPHERE
      .P1 = 16
      .P2 = 8
      Call .SetScale(0.256, 0.256, 0.256)                 '2560KM
      Call .SetPosition(122.2, 0, 0)                      'Position it Here
      Call .SetOrbitalTilt(SA_EQ + 0.3, 0, 0)             'Inclined from Centre
      Call .SetSpinDelta(0, SA_SPIN, 0)                   'Spin about y axis by these degrees each frame
      Call .SetRotation(0, 190, 0)                        '183 Abitrary
      Call .SetRotationDelta(0, -360 / 38280, 0)          '15.95d = 2400*15.95 = 38280 Frames
      Call .AttachTexture(0, 2)
    End With
    Call .AddFigure(Figure)

  End With

  'Go into a Render Loop, End It with ESC, or on FrameCounter
  Call RenderLoop(Scene, AutoIncrement:=True, FrameCounter:=4500)
  Set Scene = Nothing

End Sub

'=========================================================================================================
'=========================== THE DISPLAY RENDER LOOP CONTROL ROUTINE ========================================
'=========================================================================================================

'A Render Loop, End It with ESC - This Demo is completely controlled by program
Private Sub RenderLoop(ByRef Scene As cScene, _
                       Optional ByVal AutoIncrement As Boolean = False, _
                       Optional FrameCounter As Long = -1)

  Do
    Call Scene.Render(AutoIncrement)

    If Scene.KeyHit(vbKeyEscape) = vbKeyEscape Then Exit Do  'Check this Set of Keys (first found is returned)

    If FrameCounter > 0 Then                                 'If specified , decrement Framecounter and exit when zero
      FrameCounter = FrameCounter - 1
      If FrameCounter = 0 Then Exit Do
    End If
  Loop

End Sub

':) Ulli's VB Code Formatter V2.13.5 (31-Jan-03 12:20:23) 1 + 209 = 210 Lines
