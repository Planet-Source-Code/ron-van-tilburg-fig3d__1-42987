VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cXGen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'RVTDX.cXGen.cls - a component of the rvtDX.dll
'©2003 Ron van Tilburg - rivit@f1.net.au
'Freeware for Educational Purposes, For commercial interests contact author please, I retain copyright.

'This should be either added to rvtDX.dll or in your own program.
'If the latter the program will need a reference to DirectX 8

'Used to Generate Special Figures for the Engine
'This is based on The Template Class - You Build the GenFigure Routine

'Private Type ZVertex      'CUSTOM
'  x  As Single
'  y  As Single
'  z  As Single
'  nx As Single
'  ny As Single
'  nz As Single
'  dc As Long
'  ut As Single
'  vt As Single
'End Type

'Private Const VertexFMT As Long = D3DFVF_VERTEX Or D3DFVF_DIFFUSE

Private Const P1 As Single = 1
Private Const PH As Single = 0.5
Private Const Z0 As Single = 0
Private Const NH As Single = -0.5
Private Const N1 As Single = -1

Private NIx         As Integer    'Nr of Indices
Private zPrimitive  As Long

Public Function Generate(ByVal FigSpec As Long, _
                         ByRef zVx() As ZVertex, zIx() As Integer, _
                         ByRef Primitive As Long) As Boolean          'True is OK

'We make up a stock Object in the passed in  Vertex and Index Buffers on the given Device
'after generation we make up a VertexBuffer and IndexBuffer and Kill off the temps
'The Gen Call redimensions to zVx and zIX to the right sizes

'NOTE: Previous Vertex and IndexBuffers are FREED HERE (whether successful or not)

 Dim FT As Long, Parm1 As Long, Parm2 As Long, Parm3 As Long

  FT = FigSpec And FIGS_TYPEMASK
  Parm1 = (FigSpec And FIGS_P1MASK) / 256&
  Parm2 = (FigSpec And FIGS_P2MASK) / 65536
  Parm3 = (FigSpec And FIGS_P3MASK) / 16777216
  If Parm3 < 0 Then Parm3 = Parm3 + 256

  zPrimitive = D3DPT_TRIANGLELIST                                                      'assume this

  Select Case FT
  Case 0:  Call GenFigure
  End Select

  Primitive = zPrimitive
  NIx = 0
  Generate = True

End Function

'You build this routine to make the vertices and vertex indices (see Template for examples)
Private Sub GenFigure()

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

':) Ulli's VB Code Formatter V2.13.5 (01-Feb-03 20:49:35) 34 + 87 = 121 Lines
