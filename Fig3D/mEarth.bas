Attribute VB_Name = "mEarth"
Option Explicit

'Fig3D - a Demonstration of the capabiltites of the rvtDX.dll  D3D graphics engine
'©2003 Ron van Tilburg - rivit@f1.net.au
'Freeware for Educational Purposes, For commercial interests contact author please, I retain copyright.

'mEarth.bas       Earth-Moon System
'This demonstrates the general approach to animated scene setup

'In this Demo we follow the earth in its path around the sun in 365.25 days. During this time we rotate every
'24 hrs. The Moon too is doing its thing, once in 28 days or so. From our viewpoint we are about 500,000km
'from the earth in an orbit only slightly different from the Moon's, but only around 4 times a year. This
'means we are not far from the Moon when it crosses our viewpoint.

'Because we are a moving position in space the view of the mechanics are quite odd and take some
'getting used to. We billboard a sphere of stars around ourselves to give the impression of being in the sky.

'Nothing overly fancy, but we show how to use the Figure.Parent Property, and a point light
'Try Drawing it on Paper to see what is going on

Public Sub ShowEarth()

'TimeScale              '@50Hz
 
 Const EARTH_SPIN As Single = -10                   '1deg/Frame=36 Frames   '0.72s
 Const EARTH_ROT  As Single = EARTH_SPIN / 365.25   '13149 Frames           '4m 22s
 Const MOON_ROT   As Single = EARTH_ROT * 13.176    '  998 Frames           '20s

 Dim Scene          As cScene         'Where we will Place our Figures
 Dim Earth          As cFigure        'a useful helper
 Dim Figure         As cFigure        'a useful helper
 Dim Light          As cLight

  Set Scene = New cScene              'Create a scene (always good fun)
  With Scene                          'Create Figures in it

    .Title = "Earth-Moon seen from near the Earth as it goes around its Orbit"
    Call .AddAxes(FIGS_AXES)  'Show the Axes
    Call .Axes.SetScale(1000, 1000, 1000)
    .ShowHUD = True

    'Figures
    Call .TextureFromFile(App.Path & "\textures\Earth.bmp")       '0
    Call .TextureFromFile(App.Path & "\textures\EarthMoon.bmp")   '1

    Set Earth = New cFigure                                    'Earth
    With Earth
      .FigSpec = FIGS_SPHEROID
      .P1 = 20
      .P2 = 10
      Call .SetScale(12.756, 12.756, 12.756)
      Call .SetSpinDelta(0, EARTH_SPIN, 0)                     'Rotation about Poles
      Call .SetPosition(1000, 0, 0)
      Call .SetRotationDelta(0, EARTH_ROT, 0)                  'Speed Around The Sun
      Call .SetOrbitalTilt(23.45, 0, 0)
      Call .AttachTexture(0, 0)
    End With
    Call .AddFigure(Earth)

    Set Figure = New cFigure                                   'Moon (defined in its own Right)
    With Figure
      .FigSpec = FIGS_SPHEROID
      .P1 = 20
      .P2 = 10
      Call .SetScale(3.476, 3.476, 3.476)                      'Radii
      Call .SetSpinDelta(0, MOON_ROT, 0)                       'Rotation about Poles
      Call .SetPosition(400, 0, 0)                             'Relative To Earth
      Call .SetRotationDelta(0, MOON_ROT, 0)                   'Speed Around The Earth
      Call .SetOrbitalTilt(5.145, 0, 0)                        'Relative to Earth
      Call .AttachTexture(0, 1)
      Set .Parent = Earth                                      'Make The Linkage (always link to an earlier Figure)
    End With
    Call .AddFigure(Figure)

    Set Figure = New cFigure                                   'The Path Of Orbit
    With Figure
      .FigSpec = FIGS_CIRCLE
      .P1 = 180
      Call .SetScale(1000, 0, 1000)
      Call .SetTilt(23.45, 0, 0)
      Call .SetColourDiffuse(0.1, 0.2, 0.3)
    End With
    Call .AddFigure(Figure)

    Set Figure = New cFigure                                   'The Background Stars
    With Figure
      .FigSpec = FIGS_POINTSPHERE
      .P1 = 100
      Call .SetScale(2000, 2000, 2000)
      Call .SetColouremissive(8, 8, 8)
    End With
    Call .AddFigure(Figure)

    'Camera
    With .Camera
      Call .SetPositionRAE(500, 0, 0)
      Call .SetRotationDelta(0, EARTH_ROT * 4, 0)              'once a quarter
      Call .SetOrbitalTilt(7.145, 0, 0)                        'Relative to Earth 2deg either side of moon
      Set .Parent = Earth
    End With

    With .Viewpoint
      Set .Parent = Earth                                      'Make The Linkage (always link to an earlier Figure)
    End With

    'Light
    Call .SetAmbientLight(96, 96, 96)
    Set Light = New cLight
    With Light
      Call .MakePointlight(0, 0, 0, A0:=1, A1:=0.0001, Range:=5000)
      Call .SetColourDiffuse(8, 8, 8)
    End With
    Call .AddLight(Light)
  End With

  'Action : Go into a Render Loop, End It with ESC
  Call RenderLoop(Scene, AutoIncrement:=True)
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

':) Ulli's VB Code Formatter V2.13.5 (31-Jan-03 12:20:20) 1 + 138 = 139 Lines
