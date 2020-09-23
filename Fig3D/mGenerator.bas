Attribute VB_Name = "mGenerator"
Option Explicit

'cRVTDX.mGenerator - a component of the rvtDX.dll
'©2003 Ron van Tilburg - rivit@f1.net.au
'Freeware for Educational Purposes, For commercial interests contact author please, I retain copyright.

'Used to Generate StockObjects for the Engine
Public Const VertexFMT As Long = D3DFVF_VERTEX Or D3DFVF_DIFFUSE      'Our Vertex Format (see cDX8 for Type)

Private Const P1 As Single = 1
Private Const PH As Single = 0.5
Private Const Z0 As Single = 0
Private Const NH As Single = -0.5
Private Const N1 As Single = -1

Private NIx       As Long                       'Nr of Indices
Public Primitive  As Long                       'Filled In By Generator
Public zVx()      As ZVertex                    'Filled In By Generator
Public zIx()      As Integer                    'Filled In by Generator

Public Function GenerateFigure(ByVal FigSpec As Long) As Boolean          'True is OK

'We make up a stock Object in the passed in  Vertex and Index Buffers which are returned to make buffers
'The Gen Call redimensions to zVx and zIX to the right sizes

 Dim FT As Long, Parm1 As Long, Parm2 As Long, Parm3 As Long

  FT = FigSpec And FIGS_TYPEMASK
  Parm1 = (FigSpec And FIGS_P1MASK) / 256&
  Parm2 = (FigSpec And FIGS_P2MASK) / 65536
  Parm3 = (FigSpec And FIGS_P3MASK) / 16777216
  If Parm3 < 0 Then Parm3 = Parm3 + 256

  Primitive = D3DPT_TRIANGLELIST                                    'assume this

  'All objects are restricted to 65535 (or the max your Card can handle) vertices or Indexes
  ' this occurs when tring to put it in the VertexBuffer
  'So if youre going big and you dont see it rendered this is why
  Select Case FT
    Case FIGS_AXES:         Call GenAxes(Parm1)   'Axes P1=0, Axisplanes P1=1
  
    Case FIGS_TEXT:         'Do nothing this should not be passed in Here (Generated In Figure.GenerateMesh)
    Case FIGS_TEAPOT:       'Do nothing this should not be passed in Here (Generated In Figure.GenerateMesh)
    
    Case FIGS_POINT:        Call GenPoint
    Case FIGS_LINE:         Call GenLine
    Case FIGS_POLYLINE:     Call GenPolyLine(Parm1, 360)
    Case FIGS_POLYARC:      Call GenPolyLine(Parm1, (360 * Parm2) / 256)
  
    Case FIGS_POLYGON:      Call GenPolygon(Parm1, 360)
    Case FIGS_POLYWEDGE:    Call GenPolygon(Parm1, (360 * Parm2) / 256)     'Part of a Polygon
    Case FIGS_POLYWASHER:   Call GenPolyWasher(Parm1, Parm2 / 256)
    Case FIGS_POLYSTAR:     Call GenPolyStar(Parm1, Parm2 / 256)
    Case FIGS_POLYMESH:     Call GenPolyMesh(Parm1, Parm2)
  
    Case FIGS_REGULARSOLID: Call GenRegularSolid(Parm1)
    Case FIGS_SPHEROID:     Call GenSpheroid(Parm1, Parm2)
    Case FIGS_PRISM:        Call GenPrism(Parm1, 1)
    Case FIGS_FRUSTRUM:     Call GenPrism(Parm1, Parm2 / 256)               'Cone P2=0
    Case FIGS_TOROID:       Call GenToroid(Parm1, Parm2, Parm3 / 256)
    Case FIGS_PIPE:         Call GenPipe(Parm1, Parm2 / 256)
    Case FIGS_ASTROID:      Call GenAstroid(Parm1, Parm2)
    
    Case FIGS_NEBULA:       Call GenRotFig(1, Parm1)                        'Specials
    Case FIGS_SNAIL:        Call GenRotFig(2, Parm1)
    Case FIGS_BEAD:         Call GenRotFig(3, Parm1)
    Case FIGS_POPE:         Call GenRotFig(4, Parm1)
    Case FIGS_WALNUT:       Call GenRotFig(5, Parm1)
    Case FIGS_PERTSPHERE:   Call GenRotFig(6, Parm1)
    Case FIGS_FLOWER:       Call GenRotFig(7, Parm1)
    Case FIGS_SHELL:        Call GenRotFig(8, Parm1)
    Case FIGS_LILY:         Call GenRotFig(9, Parm1)
  
    Case FIGS_POINTFIELD:   Call GenPointField(Parm1 * 100)    'Nr of Points
    Case FIGS_POINTSPHERE:  Call GenPointSphere(Parm1 * 100)   'Nr of Points
    Case FIGS_SHEET:        Call GenSheet(Parm1, Parm2, Parm3 / 256)
    Case Else:              Call GenCube            'The Default
  End Select

  GenerateFigure = True
  NIx = 0

End Function

Public Sub EraseWorkArrays()    'Cleanup ehen no longer needed

  Erase zVx(), zIx()

End Sub

'============================== GENERATION OF STOCK OBJECTS ================================================

Private Sub GenAxes(ByVal Style As Long)
 
 Const QF As Single = 1 / 100
 Const PF As Single = 1 / 250
 Const NF As Single = -PF
 
 Const ARed As Long = &HBB7777
 Const AGreen As Long = &H77BB77
 Const ABlue As Long = &H7777BB
 
 Dim i As Long, j As Long, k As Long

  Primitive = D3DPT_LINELIST
  ReDim zVx(0 To 39), zIx(0 To 39)
  
  'x axis
  Call AddVx(0, N1, Z0, Z0, Z0, Z0, Z0, Z0, Z0, ARed)
  Call AddVx(1, P1, Z0, Z0, Z0, Z0, Z0, P1, P1, ARed)
  'axis arrow
  Call AddVx(2, P1 + QF + QF, Z0, Z0, Z0, Z0, Z0, P1, P1, ARed)
  Call AddVx(3, P1, Z0, -QF, Z0, Z0, Z0, P1, P1, ARed)
  Call AddVx(4, P1, Z0, -QF, Z0, Z0, Z0, P1, P1, ARed)
  Call AddVx(5, P1, Z0, QF, Z0, Z0, Z0, P1, P1, ARed)
  Call AddVx(6, P1, Z0, QF, Z0, Z0, Z0, P1, P1, ARed)
  Call AddVx(7, P1 + QF + QF, Z0, Z0, Z0, Z0, Z0, P1, P1, ARed)
  'axis label in XZ
  Call AddVx(8, P1 + 3 * QF, Z0, -QF, Z0, Z0, Z0, P1, P1, ARed)
  Call AddVx(9, P1 + 5 * QF, Z0, QF, Z0, Z0, Z0, P1, P1, ARed)
  Call AddVx(10, P1 + 3 * QF, Z0, QF, Z0, Z0, Z0, P1, P1, ARed)
  Call AddVx(11, P1 + 5 * QF, Z0, -QF, Z0, Z0, Z0, P1, P1, ARed)
  
  'y axis
  Call AddVx(12, Z0, P1, Z0, Z0, Z0, Z0, P1, P1, AGreen)
  Call AddVx(13, Z0, N1, Z0, Z0, Z0, Z0, Z0, Z0, AGreen)
  'axis arrow
  Call AddVx(14, Z0, P1 + QF + QF, Z0, Z0, Z0, Z0, Z0, Z0, AGreen)
  Call AddVx(15, Z0, P1, -QF, Z0, Z0, Z0, P1, P1, AGreen)
  Call AddVx(16, Z0, P1, -QF, Z0, Z0, Z0, Z0, Z0, AGreen)
  Call AddVx(17, Z0, P1, QF, Z0, Z0, Z0, P1, P1, AGreen)
  Call AddVx(18, Z0, P1, QF, Z0, Z0, Z0, Z0, Z0, AGreen)
  Call AddVx(19, Z0, P1 + QF + QF, Z0, Z0, Z0, Z0, P1, P1, AGreen)
  'axis label YZ
  Call AddVx(20, Z0, P1 + 3 * QF, Z0, Z0, Z0, Z0, P1, P1, AGreen)
  Call AddVx(21, Z0, P1 + 4 * QF, Z0, Z0, Z0, Z0, P1, P1, AGreen)
  Call AddVx(22, Z0, P1 + 5 * QF, -QF, Z0, Z0, Z0, P1, P1, AGreen)
  Call AddVx(23, Z0, P1 + 4 * QF, Z0, Z0, Z0, Z0, P1, P1, AGreen)
  Call AddVx(24, Z0, P1 + 5 * QF, QF, Z0, Z0, Z0, P1, P1, AGreen)
  Call AddVx(25, Z0, P1 + 4 * QF, Z0, Z0, Z0, Z0, P1, P1, AGreen)
  
  'z axis in XZ
  Call AddVx(26, Z0, Z0, N1, Z0, Z0, Z0, Z0, Z0, ABlue)
  Call AddVx(27, Z0, Z0, P1, Z0, Z0, Z0, P1, P1, ABlue)
  'axis arrow
  Call AddVx(28, Z0, Z0, P1 + QF + QF, Z0, Z0, Z0, Z0, Z0, ABlue)
  Call AddVx(29, -QF, Z0, P1, Z0, Z0, Z0, P1, P1, ABlue)
  Call AddVx(30, -QF, Z0, P1, Z0, Z0, Z0, Z0, Z0, ABlue)
  Call AddVx(31, QF, Z0, P1, Z0, Z0, Z0, P1, P1, ABlue)
  Call AddVx(32, QF, Z0, P1, Z0, Z0, Z0, Z0, Z0, ABlue)
  Call AddVx(33, Z0, Z0, P1 + QF + QF, Z0, Z0, Z0, P1, P1, ABlue)
  'axis label
  Call AddVx(34, -QF, Z0, P1 + 3 * QF, Z0, Z0, Z0, P1, P1, ABlue)
  Call AddVx(35, QF, Z0, P1 + 3 * QF, Z0, Z0, Z0, P1, P1, ABlue)
  Call AddVx(36, -QF, Z0, P1 + 3 * QF, Z0, Z0, Z0, P1, P1, ABlue)
  Call AddVx(37, QF, Z0, P1 + 5 * QF, Z0, Z0, Z0, P1, P1, ABlue)
  Call AddVx(38, -QF, Z0, P1 + 5 * QF, Z0, Z0, Z0, P1, P1, ABlue)
  Call AddVx(39, QF, Z0, P1 + 5 * QF, Z0, Z0, Z0, P1, P1, ABlue)

  For i = 0 To 39
    zIx(i) = i
  Next i
  NIx = 40

  If Style <> 0 Then
    ReDim Preserve zVx(0 To 1491), zIx(0 To 1491)
  
    k = 24
    For i = -5 To 5
      For j = -5 To 5
        'z planes
        Call AddVx(k + 0, j / 5 + PF, i / 5, Z0, Z0, Z0, Z0, Z0, Z0, ABlue) '0 000 00
        Call AddVx(k + 1, j / 5 + NF, i / 5, Z0, Z0, Z0, Z0, P1, P1, ABlue) '1 000 10
        Call AddVx(k + 2, i / 5 + PF, Z0, j / 5, Z0, Z0, Z0, Z0, Z0, ABlue) '0 000 00
        Call AddVx(k + 3, i / 5 + NF, Z0, j / 5, Z0, Z0, Z0, P1, P1, ABlue) '1 000 10
        'y planes
        Call AddVx(k + 4, j / 5, i / 5 + PF, Z0, Z0, Z0, Z0, Z0, Z0, AGreen) '2 000 01
        Call AddVx(k + 5, j / 5, i / 5 + NF, Z0, Z0, Z0, Z0, P1, P1, AGreen) '3 000 11
        Call AddVx(k + 6, Z0, j / 5 + PF, i / 5, Z0, Z0, Z0, Z0, Z0, AGreen) '2 000 01
        Call AddVx(k + 7, Z0, j / 5 + NF, i / 5, Z0, Z0, Z0, P1, P1, AGreen) '3 000 11
        'x planes
        Call AddVx(k + 8, i / 5, Z0, j / 5 + PF, Z0, Z0, Z0, Z0, Z0, ARed) '4 000 00
        Call AddVx(k + 9, i / 5, Z0, j / 5 + NF, Z0, Z0, Z0, P1, P1, ARed) '5 000 10
        Call AddVx(k + 10, Z0, i / 5, j / 5 + PF, Z0, Z0, Z0, Z0, Z0, ARed) '4 000 00
        Call AddVx(k + 11, Z0, i / 5, j / 5 + NF, Z0, Z0, Z0, P1, P1, ARed) '5 000 10
        k = k + 12
      Next j
    Next i
    For i = 0 To 1491
      zIx(i) = i
    Next i
    NIx = 1492
  End If
End Sub


'--------------------------- A POINT ------------------------------------------------------------------------
Private Sub GenPoint()

  Primitive = D3DPT_POINTLIST
  ReDim zVx(0 To 0), zIx(0 To 0)

  'just at 0,0,0    (this is meaningless for now)
  Call AddVx(0, 0, 0, 0, Z0, Z0, Z0, Z0, Z0, vbWhite)
  zIx(0) = 0
  NIx = 1

End Sub

'--------------------------- A POINT FIELD------------------------------------------------------------------------
Private Sub GenPointField(ByVal N As Single)    'the hundreds of points

 Dim i As Long

  Primitive = D3DPT_POINTLIST
  If N <= 0 Then N = 1000
  If N > 32000 Then N = 32000
  ReDim zVx(0 To N - 1), zIx(0 To 0)    'No IX needed, but cant be UNDEFINED

  For i = 0 To N - 1
    Call AddVx(i, Rnd() - 0.5, 0, Rnd() - 0.5, Z0, P1, Z0, Rnd, Rnd, vbWhite)
  Next i
  NIx = 0

End Sub

'--------------------------- A LINE ------------------------------------------------------------------------
Private Sub GenLine()

  Primitive = D3DPT_LINELIST
  ReDim zVx(0 To 1), zIx(0 To 1)

  'Runs from (-1,0,0 to 1,0,0)
  Call AddVx(0, N1, Z0, Z0, Z0, P1, Z0, Z0, Z0, vbWhite)
  Call AddVx(1, P1, Z0, Z0, Z0, P1, Z0, P1, P1, vbWhite)
  zIx(0) = 0
  zIx(1) = 1
  NIx = 2

End Sub

'--------------------------- A POLYLINE --------------------------------------------------------------------

Private Sub GenPolyLine(ByVal N As Long, ByVal Arc As Single)   'Parm2 is the size of Arc 0-360 degrees

 Dim u As Single, q As Single, x As Single, z As Single, ut As Single, vt As Single, da As Single
 Dim i As Long

  Primitive = D3DPT_LINELIST
  If N > 360 Then N = 360
  If Arc = 360 Then
    If N < 3 Then N = 3
    ReDim zIx(0 To 2 * N - 1) '6,8,10....
  Else
    If N < 2 Then N = 2
    ReDim zIx(0 To 2 * N - 3) '2,4,6,8...
  End If
  ReDim zVx(0 To N - 1)

  q = DtoR * Arc / N       'each step size now in radians
  If (Arc = 360) And (N And 1) = 0 Then da = q / 2             'Only for Full Polygons

  For i = 0 To N - 1
    u = q * i + da
    Call PtoR(x, z, P1, u)
    Call MapUV(ut, vt, u)
    Call AddVx(i, x, Z0, z, x, P1, z, ut, vt, UtoRGB(u))
  Next i

  For i = 0 To N - 2
    zIx(2 * i) = i
    zIx(2 * i + 1) = i + 1
    NIx = NIx + 2
  Next i

  If Arc = 360 Then            'close the circle
    zIx(2 * N - 2) = N - 1
    zIx(2 * N - 1) = 0
    NIx = NIx + 2
  End If

End Sub

'-------------------------------- POLYGONS n>=3 -----------------------------------------------------------

Private Sub GenPolygon(ByVal N As Long, ByVal Arc As Single)

 Dim u As Single, w As Single, q As Single, NVx As Long
 Dim x As Single, z As Single, ut As Single, vt As Single, da As Single
 Dim i As Long

  Primitive = D3DPT_TRIANGLELIST
  If N > 360 Then N = 360
  If Arc = 360 Then
    If N < 3 Then N = 3
  Else
    If N < 1 Then N = 1
  End If
  ReDim zVx(0 To N)   'N+1

  'a standard Polygon (or Part of One) centred at (0,0,0), radius 1
  'oriented facing UP , 0' Longitude facing x+
  'Texture Mapping set up to wrap the entire Texture centred onto the Polygon, Patches are 360/N degrees

  q = DtoR * Arc / N                                           'each step size now in radians
  If (Arc = 360) And (N And 1) = 0 Then da = q / 2             'Only for Full Polygons

  Call AddVx(0, 0, 0, 0, Z0, P1, Z0, PH, PH, vbWhite)            'The Centre
  For i = 0 To N - 1
    u = i * q + da
    Call MapUV(ut, vt, u)
    Call PtoR(x, z, P1, u)
    Call AddVx(i + 1, x, Z0, z, x, P1, z, ut, vt, UtoRGB(u))
  Next i

  For i = 1 To N - 1
    Call AddIx(0, 0, i + 1, i)  'all triangles
  Next i

  If Arc = 360 Then               'close the circle
    Call AddIx(0, 0, 1, N)
  End If

End Sub

'-------------------------------- POLYWASHER n>=3 -----------------------------------------------------------

Private Sub GenPolyWasher(ByVal N As Long, ByVal InnerR As Single)

 Dim u As Single, w As Single, q As Single, NVx As Long
 Dim x As Single, z As Single, cd As Long, da As Single
 Dim i As Long

  Primitive = D3DPT_TRIANGLELIST
  If N > 360 Then N = 360
  If N < 3 Then N = 3
  ReDim zVx(0 To 2 * N + 1) '2N+1

  'a standard Polygon Flat Ring centred at (0,0,0), radius 1
  'oriented facing UP , 0' Longitude facing x+
  'Texture Mapping set up to wrap the entire Texture centred onto the Polygon, Patches are 360/N degrees

  q = PiPI / N       'each step size now in radians
  If (N And 1) = 0 Then da = q / 2             'Only for Full Polygons

  For i = 0 To N
    u = i * q + da
    cd = UtoRGB(u)
    Call PtoR(x, z, P1, u)
    Call AddVx(2 * i, x, Z0, z, x, P1, z, i / N, Z0, cd)
    x = InnerR * x
    z = InnerR * z
    Call AddVx(2 * i + 1, x, Z0, z, x, P1, z, i / N, P1, cd)
  Next i

  For i = 0 To N - 1
    Call AddIx(2 * i, 2 * i + 2, 2 * i + 1, 2 * i + 3) 'all rectangles
  Next i
  Call AddIx(2 * N, 0, 2 * N + 1, 1)  'close the circle

End Sub

'-------------------------------- POLYSTAR n>=3 -----------------------------------------------------------

Private Sub GenPolyStar(ByVal N As Long, ByVal Pointiness As Single)

 Dim u As Single, w As Single, q As Single, NVx As Long
 Dim x As Single, z As Single, ut As Single, vt As Single
 Dim i As Long, k As Double, da As Double

  Primitive = D3DPT_TRIANGLELIST
  If N > 360 Then N = 360
  If N < 3 Then N = 3

  ReDim zVx(0 To 2 * N)  '2N+1

  'a standard Polygon Star centred at (0,0,0), radius 1
  'oriented facing UP , 0' Longitude facing x+
  'Texture Mapping set up to wrap the entire Texture centred onto the Polygon, Patches are 360/2N degrees

  q = PI / N                          'each step size now in radians
  If (N And 1) = 0 Then da = q        'offset even figures

  If N > 4 Then
    k = Sin(4# * q) / (Sin(3# * q) + Sin(q)) * (1 - Pointiness)
  Else
    k = Sin(2# * q) / Sin(q) / 2# * (1 - Pointiness)
  End If

  Call AddVx(0, 0, 0, 0, Z0, P1, Z0, 0.5, 0.5, vbWhite)                         'The Centre
  N = N + N
  For i = 0 To N - 1
    u = i * q + da
    Call MapUV(ut, vt, u)
    If (i And 1) = 0 Then
      Call PtoR(x, z, P1, u)
    Else
      Call PtoR(x, z, P1 * k, u)
    End If
    Call AddVx(i + 1, x, Z0, z, x, P1, z, ut, vt, UtoRGB(u))
  Next i

  N = N \ 2
  Call AddIx(1, 2, 2 * N, 0)             'all rectangles
  For i = 2 To N
    Call AddIx(2 * i - 1, 2 * i, 2 * i - 2, 0)
  Next i

End Sub

'-------------------------------- POLYMESH n>=3,m>0 --------------------------------------------------------

Private Sub GenPolyMesh(ByVal N As Long, ByVal M As Long)

  Primitive = D3DPT_TRIANGLELIST
  Call GenPolygon(4, 360)   'for now

End Sub

'================= We have some Classical Regular Figures nFaces = 4, 6, 8, 12, 20 =========================

Private Sub GenRegularSolid(ByVal N As Long)

  Select Case N:
  Case 4:     Call GenTetrahedron
  Case 6:     Call GenCube
  Case 8:     Call GenOctahedron
  Case 12:    Call GenDodecahedron(0)
  Case 13:    Call GenDodecahedron(2)
  Case 20:    Call GenIcosahedron
  Case 60:    Call GenDodecahedron(1)    'Not STrictly Regular
  Case Else:  Call GenCube
  End Select

End Sub

Private Sub GenTetrahedron()                '4 triangles 'Texture Needs work

 Const PQ As Single = 0.366025404      'cos 30 - 0.5
 Const NQ As Single = -PQ
 Const TV As Single = 0.067            '(1 - cos 30)/2
 Const BV As Single = 0.933            '(1 - TV)

  Primitive = D3DPT_TRIANGLELIST
  ReDim zVx(0 To 11), zIx(0 To 11)

  ' Standard Unit Tetrahedron (Side=1) centered at (0,0,0)
  '                           z+
  'Base oriented XZ          NNP(0)
  '   Facing y=DOWN          /   \
  '                    x-   / 0P0 \     x+
  '                        /  (3)  \
  '                   (1)NNN------PNN(2)
  '                           z-
  'Base
  Call AddVx(0, Z0, NQ, PH, Z0, N1, Z0, PH, TV, vbCyan)     '0 0+0 00
  Call AddVx(1, NH, NQ, NQ, Z0, N1, Z0, P1, BV, vbMagenta)  '1 0+0 01
  Call AddVx(2, PH, NQ, NQ, Z0, N1, Z0, Z0, BV, vbYellow)   '2 0+0 10
  Call AddIx(0, 0, 1, 2)

  'x+y+z+ facing
  Call AddVx(3, Z0, PH, Z0, PQ, PH, PQ, PH, TV, vbWhite)    '3 0+0 00
  Call AddVx(4, Z0, NQ, PH, PQ, PH, PQ, P1, BV, vbCyan)     '0 0+0 01
  Call AddVx(5, PH, NQ, NQ, PQ, PH, PQ, Z0, BV, vbYellow)   '2 0+0 10
  Call AddIx(3, 3, 4, 5)

  'y+z- facing
  Call AddVx(6, Z0, PH, Z0, Z0, PH, P1, PH, TV, vbWhite)    '3 0+0 00
  Call AddVx(7, PH, NQ, NQ, Z0, PH, P1, P1, BV, vbYellow)   '2 0+0 01
  Call AddVx(8, NH, NQ, NQ, Z0, PH, P1, Z0, BV, vbMagenta)  '1 0+0 10
  Call AddIx(6, 6, 7, 8)

  'x-y+z+ facing
  Call AddVx(9, Z0, PH, Z0, NQ, PH, NQ, PH, TV, vbWhite)   '3 0+0 00
  Call AddVx(10, NH, NQ, NQ, NQ, PH, NQ, P1, BV, vbMagenta) '1 0+0 01
  Call AddVx(11, Z0, NQ, PH, NQ, PH, NQ, Z0, BV, vbCyan)    '0 0+0 10
  Call AddIx(9, 9, 10, 11)

End Sub

Private Sub GenCube()                     '6 Squares, 12 Triangles

  Primitive = D3DPT_TRIANGLELIST
  ReDim zVx(0 To 23), zIx(0 To 35)

  ' Standard Unit Cube (Side=1) centered at (0,0,0)
  '
  '                   (5)NPP+========+PPP(4)
  '                        /|       /|            Y-
  '                       / |      / |            | /Z+
  '                      /  |  PPN/  |            |/
  '               (0)NPN+========+(1)|        -X--+---X+   0------tu
  '                     |(7)+----|---+PNP(6)     /|        |      |
  '                     |  /NNP  |  /           / |        |      |
  '                     | /      | /           Z- Y-       |      |
  '                     |/       |/                       tv------1
  '               (2)NNN+========+PNN(3)
  '
  'Cube has dimensions 1/2 from each axis, All Triangles are Clockwise from TopLeft of Face
  'All Normals face out from the face at all vertices (ie. each real vertex has three normals)
  'All Textures are mapped from TopLeft to BottomRight on a face as viewed from above

  'z- face
  Call AddVx(0, NH, PH, NH, Z0, Z0, N1, Z0, Z0, vbGreen)   '0 00- 00
  Call AddVx(1, PH, PH, NH, Z0, Z0, N1, P1, Z0, vbYellow)  '1 00- 10
  Call AddVx(2, NH, NH, NH, Z0, Z0, N1, Z0, P1, vbBlack)   '2 00- 01
  Call AddVx(3, PH, NH, NH, Z0, Z0, N1, P1, P1, vbRed)     '3 00- 11
  Call AddIx(0, 1, 2, 3)

  'z+ face
  Call AddVx(4, PH, PH, PH, Z0, Z0, P1, Z0, Z0, vbWhite)   '4 00+ 00
  Call AddVx(5, NH, PH, PH, Z0, Z0, P1, P1, Z0, vbCyan)    '5 00+ 10
  Call AddVx(6, PH, NH, PH, Z0, Z0, P1, Z0, P1, vbMagenta) '6 00+ 01
  Call AddVx(7, NH, NH, PH, Z0, Z0, P1, P1, P1, vbBlue)    '7 00+ 11
  Call AddIx(4, 5, 6, 7)

  'x- face
  Call AddVx(8, NH, PH, PH, N1, Z0, Z0, Z0, Z0, vbCyan)    '5 -00 00
  Call AddVx(9, NH, PH, NH, N1, Z0, Z0, P1, Z0, vbGreen)   '0 -00 10
  Call AddVx(10, NH, NH, PH, N1, Z0, Z0, Z0, P1, vbBlue)    '7 -00 01
  Call AddVx(11, NH, NH, NH, N1, Z0, Z0, P1, P1, vbBlack)   '2 -00 11
  Call AddIx(8, 9, 10, 11)

  'x+ face
  Call AddVx(12, PH, PH, NH, P1, Z0, Z0, Z0, Z0, vbYellow)  '1 +00 00
  Call AddVx(13, PH, PH, PH, P1, Z0, Z0, P1, Z0, vbWhite)   '4 +00 10
  Call AddVx(14, PH, NH, NH, P1, Z0, Z0, Z0, P1, vbRed)     '3 +00 01
  Call AddVx(15, PH, NH, PH, P1, Z0, Z0, P1, P1, vbMagenta) '6 +00 11
  Call AddIx(12, 13, 14, 15)

  'y+ face Top
  Call AddVx(16, NH, PH, PH, Z0, P1, Z0, Z0, Z0, vbCyan)    '5 0+0 11
  Call AddVx(17, PH, PH, PH, Z0, P1, Z0, P1, Z0, vbWhite)   '4 0+0 01
  Call AddVx(18, NH, PH, NH, Z0, P1, Z0, Z0, P1, vbGreen)   '0 0+0 10
  Call AddVx(19, PH, PH, NH, Z0, P1, Z0, P1, P1, vbYellow)  '1 0+0 00
  Call AddIx(16, 17, 18, 19)

  'y- face Bottom
  Call AddVx(20, PH, NH, NH, Z0, N1, Z0, Z0, Z0, vbRed)     '3 0-0 10
  Call AddVx(21, PH, NH, PH, Z0, N1, Z0, P1, Z0, vbMagenta) '6 0-0 11
  Call AddVx(22, NH, NH, NH, Z0, N1, Z0, Z0, P1, vbBlack)   '2 0-0 00
  Call AddVx(23, NH, NH, PH, Z0, N1, Z0, P1, P1, vbBlue)    '7 0-0 01
  Call AddIx(20, 21, 22, 23)

End Sub

Private Sub GenOctahedron()     '6 vertices, 24 Triangles

  Primitive = D3DPT_TRIANGLELIST
  ReDim zVx(0 To 23), zIx(0 To 23)

  ' Standard Unit Octahedron (Side=1) centered at (0,0,0)
  'ABOVE
  'x+y+z+
  Call AddVx(0, Z0, PH, Z0, PH, PH, PH, Z0, Z0, vbYellow)   '0 +++ 00
  Call AddVx(1, Z0, Z0, PH, PH, PH, PH, P1, Z0, vbCyan)     '1 +++ 10
  Call AddVx(2, PH, Z0, Z0, PH, PH, PH, Z0, P1, vbMagenta)  '2 +++ 01
  Call AddIx(0, 0, 1, 2)

  'x+y+z-
  Call AddVx(3, Z0, PH, Z0, PH, PH, NH, Z0, Z0, vbYellow)   '0 ++- 00
  Call AddVx(4, PH, Z0, Z0, PH, PH, NH, P1, Z0, vbMagenta)  '2 ++- 10
  Call AddVx(5, Z0, Z0, NH, PH, PH, NH, Z0, P1, vbRed)      '3 ++- 01
  Call AddIx(3, 3, 4, 5)

  'x-y+z-
  Call AddVx(6, Z0, PH, Z0, NH, PH, PH, Z0, Z0, vbYellow)   '0 -+- 00
  Call AddVx(7, Z0, Z0, NH, NH, PH, PH, P1, Z0, vbRed)      '3 -+- 10
  Call AddVx(8, NH, Z0, Z0, NH, PH, PH, Z0, P1, vbGreen)    '4 -+- 01
  Call AddIx(6, 6, 7, 8)

  'x-y+z+
  Call AddVx(9, Z0, PH, Z0, NH, PH, PH, Z0, Z0, vbYellow)  '0 -++ 00
  Call AddVx(10, NH, Z0, Z0, NH, PH, PH, P1, Z0, vbGreen)   '4 -++ 10
  Call AddVx(11, Z0, Z0, PH, NH, PH, PH, Z0, P1, vbCyan)    '1 -++ 01
  Call AddIx(9, 9, 10, 11)

  'BELOW
  'x+y-z+
  Call AddVx(12, Z0, NH, Z0, PH, NH, PH, P1, P1, vbBlue)    '0 +-+ 00
  Call AddVx(13, PH, Z0, Z0, PH, NH, PH, Z0, P1, vbMagenta) '2 +-+ 01
  Call AddVx(14, Z0, Z0, PH, PH, NH, PH, P1, Z0, vbCyan)    '1 +-+ 10
  Call AddIx(12, 12, 13, 14)

  'x+y-z-
  Call AddVx(15, Z0, NH, Z0, PH, NH, NH, P1, P1, vbBlue)    '0 +-- 00
  Call AddVx(16, Z0, Z0, NH, PH, NH, NH, Z0, P1, vbRed)     '3 +-- 01
  Call AddVx(17, PH, Z0, Z0, PH, NH, NH, P1, Z0, vbMagenta) '2 +-- 10
  Call AddIx(15, 15, 16, 17)

  'x-y-z-
  Call AddVx(18, Z0, NH, Z0, NH, NH, PH, P1, P1, vbBlue)    '0 --- 00
  Call AddVx(19, NH, Z0, Z0, NH, NH, PH, Z0, P1, vbGreen)   '4 --- 01
  Call AddVx(20, Z0, Z0, NH, NH, NH, PH, P1, Z0, vbRed)     '3 --- 10
  Call AddIx(18, 18, 19, 20)

  'x-y-z+
  Call AddVx(21, Z0, NH, Z0, NH, NH, PH, P1, P1, vbBlue)    '0 --+ 00
  Call AddVx(22, Z0, Z0, PH, NH, NH, PH, Z0, P1, vbCyan)    '1 --+ 01
  Call AddVx(23, NH, Z0, Z0, NH, NH, PH, P1, Z0, vbGreen)   '4 --+ 10
  Call AddIx(21, 21, 22, 23)

End Sub

Private Sub GenDodecahedron(ByVal Style As Integer)   'Twelve pentagonal faces = 60 triangles
 
 Dim u As Single, w As Single, p As Single, q As Single, NVx As Long, R2 As Single
 Dim x As Single, y As Single, z As Single, ut As Single, vt As Single
 Dim i As Long, j As Long, k As Long, M As Long

  Primitive = D3DPT_TRIANGLELIST

  ReDim zVx(0 To 31)  '12 * 16 - 1)   '=352 = 12 pentagonal faces with 3 vertices at joint edges = 12*5*3+12

  If Style = 0 Then       'Standard
    R2 = 0.79465
  ElseIf Style = 2 Then   'Recessed
    R2 = 0.58931
  Else
    R2 = P1               'Outward = 60
  End If
  
  Call AddVx(0, Z0, R2, Z0, Z0, P1, Z0, PH, PH, vbWhite)     'Top
  NVx = 1
  p = 0
  q = DtoR * 54
  For i = 0 To 4                                              'Top corners
    Call MapUV(ut, vt, p)
    Call SphtoRect(x, y, z, P1, p, q)
    Call AddVx(NVx, x, y, z, x, y, z, ut, vt, UWtoRGB(p, q))
    p = p + PiPI / 5
    NVx = NVx + 1
  Next i
  
  p = DtoR * 36
  q = DtoR * 27
  For i = 0 To 4                                              'Top centres
    Call MapUV(ut, vt, p)
    Call SphtoRect(x, y, z, R2, p, q)
    Call AddVx(NVx, x, y, z, x, y, z, PH, PH, UWtoRGB(p, q))
    p = p + PiPI / 5
    NVx = NVx + 1
  Next i
  
  p = 0
  q = DtoR * 9
  For i = 0 To 4                                              'Top centre corners
    Call MapUV(ut, vt, p)
    Call SphtoRect(x, y, z, P1, p, q)
    Call AddVx(NVx, x, y, z, x, y, z, ut, Z0, UWtoRGB(p, q))
    p = p + PiPI / 5
    NVx = NVx + 1
  Next i
  
  p = DtoR * 36
  q = -9 * DtoR
  For i = 0 To 4                                              'Bottom centre corners
    Call MapUV(ut, vt, p + DtoR * 36)
    Call SphtoRect(x, y, z, P1, p, q)
    Call AddVx(NVx, x, y, z, x, y, z, Z0, 1 - vt, UWtoRGB(p, q))
    p = p + PiPI / 5
    NVx = NVx + 1
  Next i
  
  p = 0
  q = -27 * DtoR
  For i = 0 To 4                                              'Bottom centres
    Call MapUV(ut, vt, p + DtoR * 36)
    Call SphtoRect(x, y, z, R2, p, q)
    Call AddVx(NVx, x, y, z, x, y, z, PH, PH, UWtoRGB(p, q))
    p = p + PiPI / 5
    NVx = NVx + 1
  Next i
  
  p = 36 * DtoR
  q = -54 * DtoR
  For i = 0 To 4                                              'Bottom corners
    Call MapUV(ut, vt, p + DtoR * 36)
    Call SphtoRect(x, y, z, P1, p, q)
    Call AddVx(NVx, x, y, z, x, y, z, ut, 1 - vt, UWtoRGB(p, q))
    p = p + PiPI / 5
    NVx = NVx + 1
  Next i
  
  Call AddVx(NVx, Z0, -R2, Z0, Z0, N1, Z0, PH, PH, vbWhite)     'Bottom
 
  'now do the Pentagons
  Call AddP5(0, 1, 2, 3, 4, 5)
  
  Call AddP5(6, 1, 11, 16, 12, 2)
  Call AddP5(7, 2, 12, 17, 13, 3)
  Call AddP5(8, 3, 13, 18, 14, 4)
  Call AddP5(9, 4, 14, 19, 15, 5)
  Call AddP5(10, 5, 15, 20, 11, 1)
  
  Call AddP5(21, 11, 20, 30, 26, 16)
  Call AddP5(22, 12, 16, 26, 27, 17)
  Call AddP5(23, 13, 17, 27, 28, 18)
  Call AddP5(24, 14, 18, 28, 29, 19)
  Call AddP5(25, 15, 19, 29, 30, 20)
  
  Call AddP5(31, 30, 29, 28, 27, 26)
End Sub

Private Sub AddP5(ByVal c As Integer, ByVal v1 As Integer, ByVal v2 As Integer, _
                  ByVal v3 As Integer, ByVal v4 As Integer, ByVal v5 As Integer)
  
  Call AddIx(c, c, v1, v2)
  Call AddIx(c, c, v2, v3)
  Call AddIx(c, c, v3, v4)
  Call AddIx(c, c, v4, v5)
  Call AddIx(c, c, v5, v1)
  
End Sub

Private Sub GenIcosahedron()    '20 faces, 20 triangles (6 Top, 4 Upper,4 Lower, 6 Bottom)

  Primitive = D3DPT_TRIANGLELIST
  Call GenCube    'for now

End Sub

'=========================================================================================================
Private Sub GenPointSphere(ByVal N As Single)    'the hundreds of points

 Dim i As Long
 Dim x As Single, y As Single, z As Single, u As Single, w As Single

  Primitive = D3DPT_POINTLIST
  If N <= 0 Then N = 1000
  If N > 20000 Then N = 20000
  ReDim zVx(0 To N - 1), zIx(0 To 0)    'No IX needed, but cant be ZERO

  For i = 0 To N - 1
    u = PiPI * Rnd
    w = (PI * Rnd - PI2)
    Call SphtoRect(x, y, z, P1, u, w)
    Call AddVx(i, x, y, z, x, y, z, Rnd, Rnd, UWtoRGB(u, w))
  Next i
  NIx = 0

End Sub

Private Sub GenSpheroid(ByVal NLong As Long, ByVal NLat As Long)

 Dim u As Single, w As Single, p As Single, q As Single, NVx As Long
 Dim x As Single, y As Single, z As Single, dap As Single, daq As Single
 Dim i As Long, j As Long, k As Long, M As Long

  Primitive = D3DPT_TRIANGLELIST
  If NLong < 3 Then NLong = 3
  If NLat < 2 Then NLat = 2

  k = (NLong + 1) * (NLat - 1) + 3
  ReDim zVx(0 To k)                 '(N+1)(M-1)+4

  'a standard n*m facetted sphere centred at (0,0,0)
  'oriented with North Pole facing Y+ , 0 Longitude facing x+
  'Texture Mapping set up to wrap around clockwise, Patches are 360/NLong x 180/NLat degrees in size

  Call AddVx(0, Z0, P1, Z0, Z0, P1, Z0, PH, Z0, vbWhite)     'Top    (Leading)
  Call AddVx(1, Z0, P1, Z0, Z0, P1, Z0, PH, Z0, vbWhite)     'Top    (Trailing)
  Call AddVx(2, Z0, N1, Z0, Z0, N1, Z0, PH, P1, vbWhite)     'Bottom
  Call AddVx(3, Z0, N1, Z0, Z0, N1, Z0, PH, P1, vbWhite)     'Bottom

  p = PiPI / NLong
  q = PI / NLat
  If (NLong And 1) = 0 Then dap = p / 2             'Only for Full Polygons
  If (NLat And 1) = 0 Then daq = q / 4              'color offset for even NLat
  NVx = 4
  For j = NLat - 1 To 1 Step -1                     'Generate Vertices from Nth Pole in Clockwise to Sth Pole
    w = (j - NLat / 2) * q
    For i = 0 To NLong
      u = i * p + dap
      Call SphtoRect(x, y, z, P1, u, w)
      Call AddVx(NVx, x, y, z, x, y, z, P1 - i / NLong, P1 - j / NLat, UWtoRGB(u + p / 2, w - q / 2 + daq))
      NVx = NVx + 1
    Next i
  Next j

  M = k - NLong
  For i = 1 To NLong - 1                            'Make up the Vertex Index
    Call AddIx(0, 0, 4 + i, 3 + i)                  'Triangles on Top
    Call AddIx(M + i, M + i - 1, 2, 2)              'Triangles on Bottom
  Next i
  Call AddIx(1, 1, 4 + NLong, 3 + NLong)            'Closing Top Triangle
  Call AddIx(k, k - 1, 3, 3)                        'Closing Bottom Triangle

  For j = 1 To NLat - 2
    M = 4 + (j - 1) * (NLong + 1)                         '4 ,4+ NLong+1,4+ 2NLong+2, 4+3* NLong+3...
    For i = M To M + NLong - 1
      Call AddIx(i + 1, i, i + NLong + 2, i + NLong + 1)  '5, 4, 5+(Nlong+1), 4+(Nlong+1),...
    Next i
  Next j

End Sub

'-------------------------- PRISM Rtop=1, CONOID Rtop=0, FRUSTRUM 0<RTOP<1 ----------------------------------------

Private Sub GenPrism(ByVal N As Long, ByVal RTop As Single)

 Dim u As Single, w As Single, q As Single, NVx As Long
 Dim x As Single, y As Single, z As Single, cd As Long, da As Single
 Dim ut As Single, vt As Single
 Dim i As Long

  Primitive = D3DPT_TRIANGLELIST
  If N < 3 Then N = 3
  ReDim zVx(0 To 4 * (N + 1) + 1)  '4(n+1)+2

  'a standard polygon centred at (0,0,0),extruded to radius RTop (can be 0), height 1
  'oriented on XZ plane in y+ axis , 0 Longitude facing x+
  'Texture Mapping set up to wrap the entire Texture centred over the Cone, Patches are 360/N degrees wide

  Call AddVx(0, 0, P1, 0, Z0, P1, Z0, PH, PH, vbWhite)            'The Top Centre
  Call AddVx(1, 0, Z0, 0, Z0, N1, Z0, PH, PH, vbWhite)            'The Base Centre
  NVx = 2
  q = PiPI / N
  If (N And 1) = 0 Then da = q / 2                                'Only for Full Polygons
  For i = 0 To N
    u = i * q + da
    cd = UtoRGB(u + q / 2)

    Call MapUV(ut, vt, u)
    Call PtoR(x, z, P1 * RTop, u)
    Call AddVx(NVx + 0, x, P1, z, x, P1, z, ut, vt, cd)           'Top
    Call AddVx(NVx + 1, x, P1, z, x, Z0, z, P1 - i / N, Z0, cd)   'Top Side

    Call MapUV(ut, vt, u + PI2)
    Call PtoR(x, z, P1, u)
    Call AddVx(NVx + 2, x, Z0, z, x, Z0, z, P1 - i / N, P1, cd)   'Bottom Side
    Call AddVx(NVx + 3, x, Z0, z, x, N1, z, vt, ut, cd)           'Bottom
    NVx = NVx + 4
  Next
  
  For i = 1 To N
    Call AddIx(0, 0, 4 * i + 2, 4 * i - 2)                        'Triangles TOP
    Call AddIx(4 * i + 3, 4 * i - 1, 4 * i + 4, 4 * i)            'Sides
    Call AddIx(1, 1, 4 * i + 1, 4 * i + 5)                        'Triangles BOTTOM
  Next i

End Sub

'==========================================================================================================
Private Sub GenToroid(ByVal N As Long, ByVal M As Long, ByVal InnerR As Single)

 Dim u As Single, w As Single, p As Single, q As Single, R As Single, NVx As Long
 Dim x As Single, y As Single, z As Single, cd As Long, da As Single
 Dim i As Long, j As Long

  Primitive = D3DPT_TRIANGLELIST
  If N < 3 Then N = 3
  If M < 3 Then M = 3
  ReDim zVx(0 To 4 * N * M - 1)

  'a standard n*m facetted torus centred at (0,0,0
  'oriented with North Pole facing Y+ , 0 Longitude facing x+
  'Texture Mapping set up to wrap around clockwise, Patches are 360/N x 360/M degrees in size
  'every face is individually included (hence the large number of vertices, some of which are redundant)

  p = PiPI / N
  If (N And 1) = 0 Then da = p / 2      'offset even figures
  q = PiPI / M
  R = (P1 - InnerR) / 2
  For i = 0 To N - 1
    u = i * p + da
    For j = 0 To M - 1
      w = j * q + PI

      cd = UWtoRGB(u + p / 2, w + q / 2)
      Call SphtoRectTorus(x, y, z, P1 - R, R, u, w)
      Call AddVx(NVx + 0, x, y, z, x, y, z, 1 - i / N, 1 - j / M, cd)

      Call SphtoRectTorus(x, y, z, P1 - R, R, u + p, w)
      Call AddVx(NVx + 1, x, y, z, x, y, z, 1 - (i + 1) / N, 1 - j / M, cd)

      Call SphtoRectTorus(x, y, z, P1 - R, R, u, w + q)
      Call AddVx(NVx + 2, x, y, z, x, y, z, 1 - i / N, 1 - (j + 1) / M, cd)

      Call SphtoRectTorus(x, y, z, P1 - R, R, u + p, w + q)
      Call AddVx(NVx + 3, x, y, z, x, y, z, 1 - (i + 1) / N, 1 - (j + 1) / M, cd)

      Call AddIx(NVx + 0, NVx + 1, NVx + 2, NVx + 3)    'Rectangles
      NVx = NVx + 4
    Next j
  Next i

End Sub

'Used For Toroidal Coordinates
Private Sub SphtoRectTorus(ByRef x As Single, ByRef y As Single, ByRef z As Single, _
                           ByVal R1 As Single, ByVal R2 As Single, ByVal u As Single, ByVal w As Single)

  x = (R1 + R2 * Cos(w)) * Cos(u)
  y = R2 * Sin(w)
  z = -(R1 + R2 * Cos(w)) * Sin(u)

End Sub

'==========================================================================================================
Private Sub GenPipe(ByVal N As Long, ByVal InnerR As Single)

 Dim u As Single, w As Single, q As Single, NVx As Long
 Dim x As Single, z As Single, cd As Long, da As Single
 Dim i As Long, k As Long

  Primitive = D3DPT_TRIANGLELIST
  If N > 360 Then N = 360
  If N < 3 Then N = 3
  ReDim zVx(0 To 4 * (N + 1) - 1) '4(N+1)

  'a standard Polygon Flat Ring centred at (0,0,0), radius 1
  'oriented facing UP , 0' Longitude facing x+
  'Texture Mapping set up to wrap the entire Texture centred onto the Polygon, Patches are 360/N degrees

  q = PiPI / N       'each step size now in radians
  If (N And 1) = 0 Then da = q / 2             'Only for Full Polygons

  NVx = 0
  For i = 0 To N
    u = i * q + da
    cd = UtoRGB(u)
    Call PtoR(x, z, P1, u)
    Call AddVx(NVx + 0, x, P1, z, x, P1, z, 1 - i / N, Z0, cd) 'Top    Outside
    Call AddVx(NVx + 1, x, Z0, z, x, P1, z, 1 - i / N, P1, cd) 'Bottom Outside
    x = InnerR * x
    z = InnerR * z
    Call AddVx(NVx + 2, x, P1, z, x, P1, z, i / N, Z0, cd)  'Top    Inside
    Call AddVx(NVx + 3, x, Z0, z, x, P1, z, i / N, P1, cd) 'Bottom Inside
    NVx = NVx + 4
  Next i

  For i = 0 To N - 1
    k = 4 * i
    Call AddIx(k, k + 4, k + 2, k + 6)          'TOP
    Call AddIx(k + 2, k + 4 + 2, k + 3, k + 7)  'INSIDE
    Call AddIx(k + 4, k, k + 5, k + 1)          'OUTSIDE
    Call AddIx(k + 5, k + 1, k + 7, k + 3)      'BOTTOM
  Next i

End Sub

'==========================================================================================================
Private Sub GenAstroid(ByVal NLong As Long, ByVal NLat As Long)

 Dim u As Single, w As Single, p As Single, q As Single, NVx As Long
 Dim x As Single, y As Single, z As Single, cd As Long, da As Single
 Dim i As Long, j As Long

  Primitive = D3DPT_TRIANGLELIST
  If NLong < 3 Then NLong = 3
  If NLat < 3 Then NLat = 3
  ReDim zVx(0 To 4 * NLat * NLong - 1)

  'a standard n*m facetted sphere centred at (0,0,0
  'oriented with North Pole facing Y+ , 0 Longitude facing x+
  'Texture Mapping set up to wrap around clockwise, Patches are 360/NLong x 180/NLat degrees in size
  'every face is individually included (hence the large number of vertices, some of which are redundant)

  p = PiPI / NLong
  q = PI / NLat
  If (NLong And 1) = 0 Then da = p / 2             'Only for Full Polygons
  For i = 0 To NLong - 1
    For j = 0 To NLat - 1
      u = i * p + da
      w = (j - NLat / 2) * q

      cd = UWtoRGB(u + p / 2, w + q / 2)
      Call SphtoRect(x, y, z, P1, u, w)
      x = x * x * x: y = y * y * y: z = z * z * z
      Call AddVx(NVx + 0, x, y, z, x, y, z, i / NLong, j / NLat, cd)

      Call SphtoRect(x, y, z, P1, u + p, w)
      x = x * x * x: y = y * y * y: z = z * z * z
      Call AddVx(NVx + 1, x, y, z, x, y, z, (i + 1) / NLong, j / NLat, cd)

      Call SphtoRect(x, y, z, P1, u, w + q)
      x = x * x * x: y = y * y * y: z = z * z * z
      Call AddVx(NVx + 2, x, y, z, x, y, z, i / NLong, (j + 1) / NLat, cd)

      Call SphtoRect(x, y, z, P1, u + p, w + q)
      x = x * x * x: y = y * y * y: z = z * z * z
      Call AddVx(NVx + 3, x, y, z, x, y, z, (i + 1) / NLong, (j + 1) / NLat, cd)

      If j = 0 Then
        Call AddIx(NVx + 0, NVx + 0, NVx + 2, NVx + 3)    'Triangles on Top
      ElseIf j = NLat - 1 Then
        Call AddIx(NVx + 0, NVx + 1, NVx + 2, NVx + 2)    'Triangles at Bottom
      Else
        Call AddIx(NVx + 0, NVx + 1, NVx + 2, NVx + 3)    'Rectangles
      End If
      NVx = NVx + 4
    Next j
  Next i

End Sub

'==========================================================================================================
Private Sub GenRotFig(ByVal Which As Long, ByVal N As Long)

 Dim u As Single, w As Single, p As Single, q As Single, NVx As Long
 Dim x As Single, y As Single, z As Single, dap As Single, daq As Single
 Dim i As Long, j As Long, k As Long, M As Long

  Primitive = D3DPT_TRIANGLELIST

  If N < 3 Then N = 3

  k = (N + 1) * (N - 1) + 3
  ReDim zVx(0 To k)                 '(N+1)(M-1)+4

  'a standard n*m facetted sphere centred at (0,0,0)
  'oriented with North Pole facing Y+ , 0 Longitude facing x+
  'Texture Mapping set up to wrap around clockwise, Patches are 360/NLong x 180/NLat degrees in size

  Call AddVx(0, Z0, P1, Z0, Z0, P1, Z0, PH, Z0, vbWhite)     'Top    (Leading)
  Call AddVx(1, Z0, P1, Z0, Z0, P1, Z0, PH, Z0, vbWhite)     'Top    (Trailing)
  Call AddVx(2, Z0, N1, Z0, Z0, N1, Z0, PH, P1, vbWhite)     'Bottom
  Call AddVx(3, Z0, N1, Z0, Z0, N1, Z0, PH, P1, vbWhite)     'Bottom

  p = PiPI / N
  q = PI / N
  If (N And 1) = 0 Then
    dap = p / 2           'position offset for even NLat
    daq = q / 4           'color offset for even NLat
  End If
  NVx = 4
  For j = N - 1 To 1 Step -1               'Generate Vertices from Nth Pole in Clockwise to Sth Pole
    w = (j - N / 2) * q
    For i = 0 To N
      u = i * p + dap
      Call SphtoRectRotFig(x, y, z, P1, u, w, Which)
      Call AddVx(NVx, x, y, z, x, y, z, P1 - i / N, P1 - j / N, UWtoRGB(u + p / 2, w - q / 2 + daq))
      NVx = NVx + 1
    Next i
  Next j

  M = k - N
  For i = 1 To N - 1                                    'Make up the Vertex Index
    Call AddIx(0, 0, 3 + i, 4 + i)                      'Triangles on Top
    Call AddIx(M + i, M + i - 1, 2, 2)                  'Triangles on Bottom
  Next i
  Call AddIx(1, 1, 3 + N, 4 + N)                        'Closing Top Triangle
  Call AddIx(k, k - 1, 3, 3)                            'Closing Bottom Triangle

  For j = 1 To N - 2
    M = 4 + (j - 1) * (N + 1)                           '4 ,4+ NLong+1,4+ 2NLong+2, 4+3* NLong+3...
    For i = M To M + N - 1
      Call AddIx(i + 1, i, i + N + 2, i + N + 1)        '5, 4, 5+(Nlong+1), 4+(Nlong+1),...
    Next i
  Next j

End Sub

'Used for RotationFigures (u=0 to PiPi, w=-Pi2 to Pi2)
Private Sub SphtoRectRotFig(ByRef x As Single, ByRef y As Single, ByRef z As Single, _
                            ByVal R As Single, ByVal u As Single, ByVal w As Single, ByVal Which As Long)

  Select Case Which:
    Case 0:
      x = R * Cos(u) * Cos(w)
      y = R * Sin(w)                                                           'Sphere
      z = -R * Sin(u) * Cos(w)
    Case 1:
      x = R * Cos(u) * Cos(w)
      y = R * Exp(-7 * Cos(w) * Cos(w)) * Sin(w)                               'Nebulae
      z = -R * Sin(u) * Cos(w)
    Case 2:
      R = R * Exp(-u / PiPI) * (1 + Cos(10 * u) / 20) * (1 + Cos(10 * w) / 20) 'snails
      x = R * Cos(1.1 * u) * Cos(w)
      y = R * Sin(w)
      z = -R * Sin(1.1 * u) * Cos(w)
    Case 3:
      x = R * Cos(u) * Cos(w)
      y = R * Abs(w / PI2) * Sin(w)                                            'Beads
      z = -R * Sin(u) * Cos(w)
    Case 4:
      x = R * Cos(u) * Cos(w)
      y = R * Cos(2 * u) * Sin(w)                                              'Popes
      z = -R * Sin(u) * Cos(w)
    Case 5:
      x = R * Cos(u) * Cos(w)
      y = R * Sin(w) * Sin(w) * Sin(w)                                         'Walnuts
      z = -R * Sin(u) * Cos(w)
    Case 6:
      R = R * (1 + Cos(20 * u) / 30) * (1 + Cos(10 * w) / 30)                  'Perturbed Spheres
      x = R * Cos(u) * Cos(w)
      y = R * Sin(w)
      z = -R * Sin(u) * Cos(w)
    Case 7:
      R = R * Sin(2 * (u + w)) * Sin(2 * (u - w))                              'flower heads
      x = R * Cos(u) * Cos(w)
      y = R * Sin(w)
      z = -R * Sin(u) * Cos(w)
    Case 8:
      R = R * u / PiPI                                                         'shells
      x = R * Cos(w)
      y = R * Sin(w)
      z = -R * Cos(u) * Cos(w)
    Case 9:
      x = R * u / PiPI * Cos(w)                                                'lilies
      y = R * u / PiPI * Sin(w)
      z = -R * Cos(u) * Cos(w)
  End Select
End Sub

'==========================================================================================================
Private Sub GenSheet(ByVal N As Long, ByVal M As Long, ByVal Random As Single)

 Dim u As Single, w As Single, p As Single, q As Single, NVx As Long
 Dim x As Single, y As Single, z As Single
 Dim i As Long, j As Long, k As Long

  Primitive = D3DPT_TRIANGLELIST
  If N = 0 And M = 0 Then Call GenPlane: Exit Sub
  If N < 1 Then N = 1
  If M < 1 Then M = 1
  ReDim zVx(0 To (N + 1) * (M + 1) - 1)       '(N+1)(M+1)
  
  If N = 1 Then
    p = PiPI
  ElseIf N = 2 Then
    p = PI
  Else
    p = PI2
  End If
  
  If M = 1 Then
    q = PiPI
  ElseIf M = 2 Then
    q = PI
  Else
    q = PI2
  End If
  
  NVx = 0
  For j = 0 To M
    w = j * PI2
    z = j / M - PH
    For i = 0 To N
      u = i * p
      x = i / N - PH
      y = (1 - Random) * Sin(u + w) * Sin(u - w) / 2 + Random * (2 * Rnd - 1) '-0.5 < y < 0.5
      Call AddVx(NVx, x, y, z, x, y, z, 1 - j / M, 1 - i / N, UtoRGB(PiPI * y))
      NVx = NVx + 1
    Next i
  Next j
  
  k = 0
  For j = 0 To M - 1
    For i = 0 To N - 1
      Call AddIx(k, k + 1, k + (N + 1), k + (N + 1) + 1)
      k = k + 1
    Next i
    k = k + 1
  Next j
End Sub

'==========================================================================================================

'Some Planar Figures ------------------ ie. they have no depth (so women find them unattractive)

Private Sub GenPlane()

  Primitive = D3DPT_TRIANGLELIST
  ReDim zVx(0 To 3), zIx(0 To 5)
  '                           z+
  'Plane oriented XZ   (2)N0P------P0P(0)
  '   Facing y=UP          |        |
  '                     x- |        | x+
  '                        |        |
  '                    (3)N0N------P0N(1)
  '                           z-
  Call AddVx(0, PH, Z0, PH, Z0, P1, Z0, Z0, Z0, vbWhite) '1 0+0 10
  Call AddVx(1, PH, Z0, NH, Z0, P1, Z0, P1, Z0, vbBlack) '3 0+0 11
  Call AddVx(2, NH, Z0, PH, Z0, P1, Z0, Z0, P1, vbBlack) '2 0+0 01
  Call AddVx(3, NH, Z0, NH, Z0, P1, Z0, P1, P1, vbWhite) '0 0+0 00
  Call AddIx(0, 1, 2, 3)

End Sub

Private Sub GenTriangle()

 Const PQ As Single = 0.366025404      'sin 60 - 0.5
 Const NQ As Single = -PQ
 Const TV As Single = 0.067
 Const BV As Single = 0.933

  Primitive = D3DPT_TRIANGLELIST
  ReDim zVx(0 To 2), zIx(0 To 2)

  '                           z+
  'Plane oriented XZ          00P(0)
  '   Facing y=UP            /   \
  '                    x-   /  0  \     x+
  '                        /       \
  '                   (2)N0N------P0N(1)
  '                           z-
  Call AddVx(0, Z0, Z0, PH, Z0, P1, Z0, PH, TV)   '0 0+0 00
  Call AddVx(1, PH, Z0, PQ, Z0, P1, Z0, P1, BV)   '1 0+0 10
  Call AddVx(2, NH, Z0, NQ, Z0, P1, Z0, Z0, BV)   '2 0+0 01
  Call AddIx(0, 0, 1, 2)

End Sub

'======================================== HELPERS =========================================================
Private Sub AddVx(ByVal Index As Integer, _
                  ByVal x As Single, ByVal y As Single, ByVal z As Single, _
                  ByVal nx As Single, ByVal ny As Single, ByVal nz As Single, _
                  ByVal ut As Single, ByVal vt As Single, Optional Diffuse As Long = vbWhite)

  With zVx(Index)
    .x = x
    .y = y
    .z = z
    .nx = nx
    .ny = ny
    .nz = nz
    .dc = Diffuse
    .ut = ut
    .vt = vt
  End With

End Sub

'Add entries to the vertex Index by passing the rectangle ends to it (clockwise as viewed)
Private Sub AddIx(ByVal TL As Integer, ByVal TR As Integer, ByVal BL As Integer, ByVal BR As Integer)

  If TL <> TR And BL <> BR Then
    ReDim Preserve zIx(0 To NIx + 5)
    zIx(NIx + 0) = TL
    zIx(NIx + 1) = TR
    zIx(NIx + 2) = BL
    zIx(NIx + 3) = BL
    zIx(NIx + 4) = TR
    zIx(NIx + 5) = BR
    NIx = NIx + 6
  ElseIf TL = TR Then      'An Up Triangle
    ReDim Preserve zIx(0 To NIx + 2)
    zIx(NIx + 0) = TL
    zIx(NIx + 1) = BR
    zIx(NIx + 2) = BL
    NIx = NIx + 3
  ElseIf BL = BR Then      'A Down Triangle
    ReDim Preserve zIx(0 To NIx + 2)
    zIx(NIx + 0) = TL
    zIx(NIx + 1) = TR
    zIx(NIx + 2) = BL
    NIx = NIx + 3
  Else
    'PANIC a line or point given
  End If

End Sub

':) Ulli's VB Code Formatter V2.13.5 (01-Feb-03 20:47:50) 19 + 1059 = 1078 Lines
