Attribute VB_Name = "mFigures"
Option Explicit

'Fig3D - a Demonstration of the capabiltites of the rvtDX.dll  D3D graphics engine
'©2003 Ron van Tilburg - rivit@f1.net.au
'Freeware for Educational Purposes, For commercial interests contact author please, I retain copyright.

'This demonstrates a number of the available figures and the general approach to scene setup
'It also shows some of the effects that can be created by influencing RenderMode

Public Sub ShowFigures()

 Dim Scene      As cScene            'Where we will Place our Figures
 Dim Figure     As cFigure           'a useful helper
 Dim Light      As cLight            'a useful helper
 Dim i As Long, j As Long, k As Long
 Dim TestFont As StdFont

  Set Scene = New cScene              'Create a scene (always good fun)
  With Scene                          'Create Figures in it
    Call .AddAxes(FIGS_AXES)                    'Show the Axes
    Call .Camera.SetPosition(300, 300, 0)        'View From ABove
    
    Call .TextureFromFile(App.Path & "\textures\Fig_ThisSideUp.bmp")  '0
    Call .TextureFromFile(App.Path & "\textures\Fig_Exotic.bmp")      '1
    Call .TextureFromFile(App.Path & "\textures\Fig_GreenIron.bmp")   '2
    Call .TextureFromFile(App.Path & "\textures\Fig_CyanTiles.bmp")   '3
    Call .TextureFromFile(App.Path & "\textures\Fig_Phoenix.bmp")     '4
    Call .TextureFromFile(App.Path & "\textures\Fig_Plastic.bmp")     '5
    Call .TextureFromFile(App.Path & "\textures\Earth.bmp")           '6
    Call .TextureFromFile(App.Path & "\textures\EarthMoon.bmp")       '7
    Call .TextureFromFile(App.Path & "\textures\Saturn.bmp")          '8
    Call .TextureFromFile(App.Path & "\textures\SaturnRings.bmp")     '9
    Call .TextureFromFile(App.Path & "\textures\SaturnTitan.bmp")     '10
    Call .TextureFromFile(App.Path & "\textures\StarsAndHoles.bmp")   '11
    
    Set Light = New cLight
    With Light
      Call .MakeDirectionalLight(-1, -1, -1)
    End With
    Call .AddLight(Light)                                   'Just one light

    'Now make figures (all of which will be scaled to the same size settings

    j = -110: k = -50
    For i = 1 To 9
      Set Figure = New cFigure
      With Figure
        .FigSpec = FIGS_POLYLINE                            'Plane Outlines
        .P1 = i + 2
        Call .SetPosition(j + 20 * ((i - 1) \ 3), 0, k + 20 * ((i - 1) Mod 3))
        Call .SetScale(10, 10, 10)
        Call .AttachTexture(0, 0)
      End With
      Call .AddFigure(Figure)
    Next i
    
    j = -110: k = 10
    For i = 1 To 9
      Set Figure = New cFigure
      With Figure
        .FigSpec = FIGS_POLYGON                             'Polygons
        .P1 = i + 2
        Call .SetPosition(j + 20 * ((i - 1) \ 3), 0, k + 20 * ((i - 1) Mod 3))
        Call .SetScale(10, 10, 10)
        Call .AttachTexture(0, 0)
      End With
      Call .AddFigure(Figure)
    Next i

    j = -50: k = -110
    For i = 1 To 9
      Set Figure = New cFigure
      With Figure
        .FigSpec = FIGS_POLYARC                             'Part Outlines
        .P1 = i + 1
        .P2 = 57
        Call .SetPosition(j + 20 * ((i - 1) \ 3), 0, k + 20 * ((i - 1) Mod 3))
        Call .SetScale(10, 10, 10)
        Call .AttachTexture(0, 0)
      End With
      Call .AddFigure(Figure)
    Next i
    
    j = 10: k = -110
    For i = 1 To 9
      Set Figure = New cFigure
      With Figure
        .FigSpec = FIGS_POLYWEDGE                          'Polygon Pie Wedges
        .P1 = i + 1
        .P2 = 57
        Call .SetPosition(j + 20 * ((i - 1) \ 3), 0, k + 20 * ((i - 1) Mod 3))
        Call .SetScale(10, 10, 10)
        Call .AttachTexture(0, 0)
      End With
      Call .AddFigure(Figure)
    Next i
    
    j = 10: k = 70
    For i = 1 To 9
      Set Figure = New cFigure
      With Figure
        .FigSpec = FIGS_POLYWASHER                         'Polygon Rings
        .P1 = i + 2
        .P2 = 0.5
        Call .SetPosition(j + 20 * ((i - 1) \ 3), 0, k + 20 * ((i - 1) Mod 3))
        Call .SetScale(10, 10, 10)
        Call .AttachTexture(0, 0)
      End With
      Call .AddFigure(Figure)
    Next i
    
    j = 70: k = 10
    For i = 1 To 9
      Set Figure = New cFigure
      With Figure
        .FigSpec = FIGS_POLYSTAR                           'Polygon Stars
        .P1 = i + 2
        Call .SetPosition(j + 20 * ((i - 1) \ 3), 0, k + 20 * ((i - 1) Mod 3))
        Call .SetScale(10, 10, 10)
        Call .AttachTexture(0, 0)
      End With
      Call .AddFigure(Figure)
    Next i
    
    j = 70: k = -50
    For i = 1 To 9
      Set Figure = New cFigure
      With Figure
        .FigSpec = FIGS_POLYSTAR                           'Polygon Stars but Pointier
        .P1 = i + 2
        .P2 = 0.5
        Call .SetPosition(j + 20 * ((i - 1) \ 3), 0, k + 20 * ((i - 1) Mod 3))
        Call .SetScale(10, 10, 10)
        Call .AttachTexture(0, 0)
      End With
      Call .AddFigure(Figure)
    Next i
    
    j = -110: k = 70
    For i = 1 To 5
      Set Figure = New cFigure
      With Figure
        .FigSpec = FIGS_REGULARSOLID                       'Regular Solids
        Select Case i
          Case 1: .P1 = 4
          Case 2: .P1 = 6
          Case 3: .P1 = 8
          Case 4: .P1 = 12
          Case 5: .P1 = 20
        End Select
        Call .SetPosition(j + 20 * ((i - 1) \ 3), 0, k + 20 * ((i - 1) Mod 3))
        Call .SetScale(10, 10, 10)
        Call .AttachTexture(0, 0)
      End With
      Call .AddFigure(Figure)
    Next i

    j = -110: k = 70: i = 6
    Set Figure = New cFigure
    With Figure
      .FigSpec = FIGS_POINTFIELD                            'A square of random points
      .P1 = 1
        Call .SetPosition(j + 20 * ((i - 1) \ 3), 0, k + 20 * ((i - 1) Mod 3))
      Call .SetScale(10, 10, 10)
      Call .AttachTexture(0, 0)
    End With
    Call .AddFigure(Figure)

    j = -110: k = 70: i = 7
    Set Figure = New cFigure
    With Figure
      .FigSpec = FIGS_POINTSPHERE                           'a random sphere outline
      .P1 = 1
        Call .SetPosition(j + 20 * ((i - 1) \ 3), 0, k + 20 * ((i - 1) Mod 3))
      Call .SetScale(10, 10, 10)
      Call .AttachTexture(0, 0)
    End With
    Call .AddFigure(Figure)
 
    j = -110: k = 70: i = 8
    Set Figure = New cFigure
    With Figure
      .FigSpec = FIGS_TEAPOT                                'a teapot
      Call .SetPosition(j + 20 * ((i - 1) \ 3), 0, k + 20 * ((i - 1) Mod 3))
      Call .SetScale(10, 10, 10)
      Call .AttachTexture(0, 0)
    End With
    Call .AddFigure(Figure)
    
    j = -50: k = 10
    For i = 1 To 9
      Set Figure = New cFigure
      With Figure
        .FigSpec = FIGS_SPHEROID                          'Solids inside spheres
        .P1 = i + 2
        .P2 = 3 + i / 2
        Call .SetPosition(j + 20 * ((i - 1) \ 3), 0, k + 20 * ((i - 1) Mod 3))
        Call .SetScale(10, 10, 10)
        Call .AttachTexture(0, 0)
      End With
      Call .AddFigure(Figure)
    Next i
    
    j = -50: k = 70
    For i = 1 To 9
      Set Figure = New cFigure
      With Figure
        .FigSpec = FIGS_SHEET                             'Modulated Plane
        .P1 = i
        .P2 = i + 1
        Call .SetPosition(j + 20 * ((i - 1) \ 3), 0, k + 20 * ((i - 1) Mod 3))
        Call .SetScale(10, 10, 10)
        Call .AttachTexture(0, 0)
      End With
      Call .AddFigure(Figure)
    Next i
    
    j = -170: k = -50
    For i = 1 To 9
      Set Figure = New cFigure
      With Figure
        .FigSpec = FIGS_NEBULA + i - 1                   'Some strange ones
        .P1 = 10
        Call .SetPosition(j + 20 * ((i - 1) \ 3), 0, k + 20 * ((i - 1) Mod 3))
        Call .SetScale(10, 10, 10)
        Call .AttachTexture(0, 0)
      End With
      Call .AddFigure(Figure)
    Next i
    
    j = -170: k = 10
    For i = 1 To 9
      Set Figure = New cFigure
      With Figure
        .FigSpec = FIGS_SHEET                             'Randomised Modulated Plane
        .P1 = i
        .P2 = i + 1
        .P3 = Rnd
        Call .SetPosition(j + 20 * ((i - 1) \ 3), 0, k + 20 * ((i - 1) Mod 3))
        Call .SetScale(10, 10, 10)
        Call .AttachTexture(0, 0)
      End With
      Call .AddFigure(Figure)
    Next i
    
    j = -50: k = -50
    For i = 1 To 9
      Set Figure = New cFigure
      With Figure
        .FigSpec = FIGS_PRISM                             'Prisms
        .P1 = i + 2
        Call .SetPosition(j + 20 * ((i - 1) \ 3), 0, k + 20 * ((i - 1) Mod 3))
        Call .SetScale(10, 10, 10)
        Call .AttachTexture(0, 0)
      End With
      Call .AddFigure(Figure)
    Next i
    
    j = 10: k = -50
    For i = 1 To 9
      Set Figure = New cFigure
      With Figure
        .FigSpec = FIGS_FRUSTRUM                          'Frustrums
        .P1 = i + 2
        .P2 = 0.5
        Call .SetPosition(j + 20 * ((i - 1) \ 3), 0, k + 20 * ((i - 1) Mod 3))
        Call .SetScale(10, 10, 10)
        Call .AttachTexture(0, 0)
      End With
      Call .AddFigure(Figure)
    Next i
        
    j = 10: k = 10
    For i = 1 To 9
      Set Figure = New cFigure
      With Figure
        .FigSpec = FIGS_FRUSTRUM                          'Cones = Frustrum with P2=0
        .P1 = i + 2
        Call .SetPosition(j + 20 * ((i - 1) \ 3), 0, k + 20 * ((i - 1) Mod 3))
        Call .SetScale(10, 10, 10)
        Call .AttachTexture(0, 0)
      End With
      Call .AddFigure(Figure)
    Next i
    
    j = 70: k = 70
    For i = 1 To 9
      Set Figure = New cFigure
      With Figure
        .FigSpec = FIGS_TOROID                          'Toroids
        .P1 = i + 2
        .P2 = 3 + i / 2
        .P3 = 0.5
        Call .SetPosition(j + 20 * ((i - 1) \ 3), 0, k + 20 * ((i - 1) Mod 3))
        Call .SetScale(10, 10, 10)
        Call .AttachTexture(0, 0)
      End With
      Call .AddFigure(Figure)
    Next i
    
    j = -110: k = -110
    For i = 1 To 9
      Set Figure = New cFigure
      With Figure
        .FigSpec = FIGS_PIPE                            'Pipes
        .P1 = i + 2
        .P2 = 0.5
        Call .SetPosition(j + 20 * ((i - 1) \ 3), 0, k + 20 * ((i - 1) Mod 3))
        Call .SetScale(10, 10, 10)
        Call .AttachTexture(0, 0)
      End With
      Call .AddFigure(Figure)
    Next i
   
    j = 70: k = -110
    For i = 1 To 9
      Set Figure = New cFigure
      With Figure
        .FigSpec = FIGS_ASTROID                         'Astroids
        .P1 = i + 2
        .P2 = 3 + i / 2
        Call .SetPosition(j + 20 * ((i - 1) \ 3), 0, k + 20 * ((i - 1) Mod 3))
        Call .SetScale(10, 10, 10)
        Call .AttachTexture(0, 0)
      End With
      Call .AddFigure(Figure)
    Next i
    
    Set TestFont = New StdFont
    TestFont.Name = "Courier New"
    TestFont.Size = 10
    
    j = -170: k = -110
    For i = 1 To 9
      Set Figure = New cFigure
      With Figure
        .FigSpec = FIGS_TEXT                            'Text (always faces to Z along X from 0,0,0)
        .Text = Chr$(64 + i)
        Set .TextFont = TestFont
        Call .SetPosition(j + 20 * ((i - 1) \ 3), 0, k + 20 * ((i - 1) Mod 3))
        Call .SetSpin(0, -90, 0)
        Call .SetScale(10, 10, 10)
        Call .AttachTexture(0, 0)
      End With
      Call .AddFigure(Figure)
    Next i
  End With
'GoTo zzz
'-----------------------------------------------------------------------------------------------------------
  'Go into a Render Loop, End It with ESC
  Scene.Title = "Solid Vertex Colours"
  Scene.RenderFlags = RF_VERTEXCOLOURS Or RF_NOTEXTURES Or RF_DRAWSOLID
  Call RenderLoop(Scene, FrameCounter:=500)
  
  Scene.Title = "Vertex Colours in Points only"
  Scene.RenderFlags = RF_VERTEXCOLOURS Or RF_NOTEXTURES Or RF_DRAWPOINTS
  Call RenderLoop(Scene, FrameCounter:=500)
  
  Scene.Title = "Vertex Colours In WireFrame"
  Scene.RenderFlags = RF_VERTEXCOLOURS Or RF_NOTEXTURES Or RF_DRAWWIREFRAME
  Call RenderLoop(Scene, FrameCounter:=500)
  
  Scene.Title = "Vertex Colours In WireFrame with hidden surfaces removed"
  Scene.RenderFlags = RF_VERTEXCOLOURS Or RF_NOTEXTURES Or RF_DRAWWIREFRAME Or RF_REMOVEHIDDEN
  Call RenderLoop(Scene, FrameCounter:=500)
  
'-----------------------------------------------------------------------------------------------------------
  Scene.Title = "No Vertex Colours, Default white Material"
  Scene.RenderFlags = RF_NOTEXTURES
  Call RenderLoop(Scene, FrameCounter:=500)
  
'-----------------------------------------------------------------------------------------------------------
  Scene.Title = "No Vertex Colours, but With Test Texture"
  Scene.RenderFlags = RF_DRAWSOLID
  Call RenderLoop(Scene, FrameCounter:=500)

'-----------------------------------------------------------------------------------------------------------
  Scene.Title = "No Vertex Colours, but With Random Textures"
  i = 1
  Do
    Set Figure = Scene.Figure(i)
    If Figure Is Nothing Then Exit Do
    Call Figure.AttachTexture(0, Int(Rnd * 12), RF_LIGHTTINT)
    i = i + 1
  Loop
  Scene.RenderFlags = RF_DRAWSOLID
  Call RenderLoop(Scene, FrameCounter:=500)
  
  Scene.Title = "No Vertex Colours, but With Random Textures, in Wireframe"
  Scene.RenderFlags = RF_DRAWWIREFRAME
  Call RenderLoop(Scene, FrameCounter:=500)

'-----------------------------------------------------------------------------------------------------------
  i = 1
  Do
    Set Figure = Scene.Figure(i)
    If Figure Is Nothing Then Exit Do
    Call Figure.AttachTexture(0, Int(Rnd * 12))
    i = i + 1
  Loop
  Scene.Title = "Random Textures and with DarkTint"
  Scene.RenderFlags = RF_DRAWSOLID Or RF_DARKTINT
  Call RenderLoop(Scene, FrameCounter:=500)
  
  Scene.Title = "Random Textures and with Shiny?"
  Scene.RenderFlags = RF_DRAWSOLID Or RF_SHINY
  Call RenderLoop(Scene, FrameCounter:=500)
  
  Scene.Title = "Random Textures and with Transparent"
  Scene.RenderFlags = RF_DRAWSOLID Or RF_TRANSPARENT
  Call RenderLoop(Scene, FrameCounter:=500)
  
  Scene.Title = "Random Textures and with LightTint"
  Scene.RenderFlags = RF_DRAWSOLID Or RF_LIGHTTINT
  Call RenderLoop(Scene, FrameCounter:=500)
  
zzz:
'-----------------------------------------------------------------------------------------------------------
  Scene.Title = "The Works - Everything Random - Thankyou and Goodbye"
  i = 1
  Randomize
  Do
    Set Figure = Scene.Figure(i)
    If Figure Is Nothing Then Exit Do
    With Figure
      Call .SetScale(8 + 8 * Rnd, 8 + 8 * Rnd, 8 + 8 * Rnd)
      Call .SetSpinDelta(4 * (Rnd - 0.5), 4 * (Rnd - 0.5), 4 * (Rnd - 0.5))
      Call .SetPositionDelta((Rnd - 0.5) / 10, (Rnd - 0.5) / 10, (Rnd - 0.5) / 10)
      Call .AttachTexture(0, Int(Rnd * 12))
      Call .SetColouremissive(Rnd, Rnd, Rnd, Rnd)
      .RenderFlags = (Rnd * &H1FF&) And &H1FC&                        'Random Effects
    End With
    i = i + 1
  Loop
  Scene.ShowHUD = True
  Call Scene.SetBackground(&H50, &H50, &H40)
  Call Scene.Camera.SetPosition(300, 100, 0)        'View From ABove
  Call Scene.Camera.SetRotationDelta(0, 1, 0)
  
  'Let Figures override the Scene RenderFlags
  Scene.RenderFlags = RF_DRAWSOLID
  Call RenderLoop(Scene, AutoIncrement:=True, FrameCounter:=1500)
  
  
  Set Scene = Nothing

End Sub

'=========================================================================================================
'=========================== THE DISPLAY RENDER LOOP CONTROL ROUTINE ========================================
'=========================================================================================================

'A Render Loop, End It with ESC, navigate with Keys or Mouse
Private Sub RenderLoop(ByRef Scene As cScene, _
                       Optional ByVal AutoIncrement As Boolean = False, _
                       Optional FrameCounter As Long = -1)

'These are changed with the Arrow keys and Function Keys

 Dim dRange As Single, dAzimuth As Single, dElevation As Single, UsedKeys As Boolean

  Do
    Call Scene.Render(AutoIncrement)

    UsedKeys = True                                          'Assume a key was hit
    Select Case Scene.KeyHit(vbKeyEscape, _
           vbKeyPageUp, vbKeyPageDown, _
           vbKeyInsert, vbKeyDelete, _
           vbKeyHome, vbKeyEnd, _
           vbKeyScrollLock)              'Check this Set of Keys (first found is returned)

    Case vbKeyEscape:   Exit Do                              'if escape was pressed, exit program

    Case vbKeyPageUp:   dElevation = 1                       'viewpoint up
    Case vbKeyPageDown: dElevation = -1                      'viewpoint down
    Case vbKeyHome:     dAzimuth = -1                        'viewpoint clockwise
    Case vbKeyEnd:      dAzimuth = 1                         'viewpoint anticlockwise
    Case vbKeyDelete:   dRange = 0.5                         'zoom out
    Case vbKeyInsert:   dRange = -0.5                        'zoom in

    Case vbKeyScrollLock:
      Call Scene.Viewpoint.SetPosition(0, 0, 0)              'zoom in

    Case Else:          UsedKeys = False                     'Usable key was not hit
    End Select

    If UsedKeys Then                                           'we need to Move the Camera
      Call Scene.MoveCameraRAE(dRange, dAzimuth, dElevation)
      dRange = 0: dAzimuth = 0: dElevation = 0
    Else                                                       'Move it by The mouse (if it was used)
      Call Scene.MoveViewpointMouse(0.1, 0.1, 0.01)
      'Call Scene.MoveCameraMouse(0.1, 0.1, 0.01)              'Divide the Mouse Movements by 10,10,100
    End If

    'If specified , decrement Framecounter and exit when zero
    If FrameCounter > 0 Then
      FrameCounter = FrameCounter - 1
      If FrameCounter = 0 Then Exit Do
    End If
  Loop

End Sub

':) Ulli's VB Code Formatter V2.13.5 (31-Jan-03 12:20:17) 1 + 222 = 223 Lines
