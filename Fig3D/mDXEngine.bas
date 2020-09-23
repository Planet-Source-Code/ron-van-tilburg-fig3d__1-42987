Attribute VB_Name = "mDXEngine"
Option Explicit

'cRVTDX.mDXEngine - a component of the rvtDX.dll
'©2003 Ron van Tilburg - rivit@f1.net.au
'Freeware for Educational Purposes, For commercial interests contact author please, I retain copyright.

'A (very) simple DirectX8 Graphics/Sound/Input Engine

'TO USE THIS CODE:
'You Need DirectX8 for Visual Basic Type Library, from www.microsoft.com/directx
'You need this Project to Reference "DirectX8 for Visual Basic Type Library"
'
'This File Contains All of the DX Control Routines

'The Possible Figure Types That we can recognise

' Some useful Constatnts
Public Const PI   As Double = 3.141592653589
Public Const PiPI As Double = PI + PI             '360 deg
Public Const PI2  As Double = PI / 2#             ' 90 deg
Public Const PI4  As Double = PI / 4#             ' 45 deg
Public Const DtoR As Double = PI / 180#
Public Const RtoD As Double = 180# / PI

' GLOBAL VARIABLES
Public DXhWnd         As Long                        'The Level 1 (window) Form to which we are connected
Public DXHWndRect     As RECT                        'The Size of the window we started with
Public DX8            As DirectX8                    'the DirectX8 object - The Master of Ceremonies

'Graphics Objects of DirectX
Public D3D8           As Direct3D8                   'Responsible For Graphics Production
Public D3DX           As D3DX8                       'Graphics Utilities and Maths
Public D3DD           As Direct3DDevice8             'The Rendering Device We will Use
Public D3DDC          As D3DCAPS8                    'The Device Capabilities available
Public D3DDM          As D3DDISPLAYMODE              'The DisplayMode in Effect

'Sound Objects
Public DS8            As DirectSound8                'For making noises
Public Sound          As DirectSoundSecondaryBuffer8 'Somewhere to Put A WAV
Public SoundDesc      As DSBUFFERDESC                'And its Description

'Input Devices at our Disposal
Public DI8            As DirectInput8                'used to get data from input from the mouse and keyboard.

Public Keyboard       As DirectInputDevice8
Public KeyboardState  As DIKEYBOARDSTATE             'Keyboard state data

Public Mouse          As DirectInputDevice8
Public MouseState     As DIMOUSESTATE                'Mouse State Data

'Some Useful Constant Matrices
Public WORLD_ORIGIN   As D3DVECTOR                   'Nominally (0,0,0)
Public WORLD_YISUP    As D3DVECTOR                   'Usually   (0,1,0)


'Used to do fontprocessing
Private Declare Function GetDC Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32" (ByVal hwnd As Long, ByVal hDC As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
'IN cDX8
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

'================== MAJOR TRANSFORMATION MATRICES +========================================================

Public Sub MakeFigureMatrix(ByRef XForm As D3DMATRIX, _
                            ByRef Size As D3DVECTOR, _
                            ByRef Spin As D3DVECTOR, _
                            ByRef Tilt As D3DVECTOR)

 Dim tmpM As D3DMATRIX

  Call D3DXMatrixIdentity(XForm)                                                   'identity
  'Scale
  Call D3DXMatrixScaling(tmpM, Size.x, Size.y, Size.z)
  Call D3DXMatrixMultiply(XForm, XForm, tmpM)                                      'Scaled
  'Rotate
  Call D3DXMatrixRotationYawPitchRoll(tmpM, Spin.y, Spin.x, Spin.z)                'Spin   about Axes
  Call D3DXMatrixMultiply(XForm, XForm, tmpM)
  'Tilt
  Call D3DXMatrixRotationYawPitchRoll(tmpM, Tilt.y, Tilt.x, Tilt.z)                'Tilted about Axes
  Call D3DXMatrixMultiply(XForm, XForm, tmpM)

End Sub

Public Sub MakeOrbitMatrix(ByRef XForm As D3DMATRIX, _
                           ByRef Position As D3DVECTOR, _
                           ByRef Rotation As D3DVECTOR, _
                           ByRef Tilt As D3DVECTOR, _
                           ByRef Origin As D3DVECTOR)

 Dim tmpM As D3DMATRIX

  Call D3DXMatrixIdentity(XForm)                                         'identity
  'Translate to Orbital Position
  Call D3DXMatrixTranslation(tmpM, Position.x, Position.y, Position.z)             'to Orbital XYZ
  Call D3DXMatrixMultiply(XForm, XForm, tmpM)                                      'Translated
  'Rotate around Orbit
  Call D3DXMatrixRotationYawPitchRoll(tmpM, Rotation.y, Rotation.x, Rotation.z)    'Rotated about Axes
  Call D3DXMatrixMultiply(XForm, XForm, tmpM)
  'Tilt Orbit
  Call D3DXMatrixRotationYawPitchRoll(tmpM, Tilt.y, Tilt.x, Tilt.z)                 'Tilted about Axes
  Call D3DXMatrixMultiply(XForm, XForm, tmpM)
  'Translate Orbit to Orbital Origin
  Call D3DXMatrixTranslation(tmpM, Origin.x, Origin.y, Origin.z)                   'to Orbital Origin
  Call D3DXMatrixMultiply(XForm, XForm, tmpM)                                      'Translated

End Sub

'============================ VECTOR OPS ====================================================================
Public Sub SetVector(ByRef Vec As D3DVECTOR, ByVal x As Single, ByVal y As Single, ByVal z As Single)

  With Vec
    .x = x: .y = y: .z = z
  End With

End Sub

Public Sub IncrementVector(ByRef Vec As D3DVECTOR, ByRef dVec As D3DVECTOR)

  Vec.x = Vec.x + dVec.x
  Vec.y = Vec.y + dVec.y
  Vec.z = Vec.z + dVec.z

End Sub

Public Function VectorDiff(ByRef Vec As D3DVECTOR, ByRef dVec As D3DVECTOR) As D3DVECTOR

  VectorDiff.x = Vec.x - dVec.x
  VectorDiff.y = Vec.y - dVec.y
  VectorDiff.z = Vec.z - dVec.z

End Function

Public Function IsZeroVector(ByRef Vec As D3DVECTOR) As Boolean

  If Vec.x = 0 Then
    If Vec.y = 0 Then
      If Vec.z = 0 Then
        IsZeroVector = True
      End If
    End If
  End If

End Function

'============================ COORDINATE  OPS ==============================================================
'From Spherical Coordinates to Rectangular 3D coordinates (LH System)   (RADIANS)
Public Sub SphtoRect(ByRef x As Single, ByRef y As Single, ByRef z As Single, _
                     ByVal Range As Single, ByVal Azimuth As Single, ByVal Elevation As Single)

  x = Range * Cos(Azimuth) * Cos(Elevation)
  y = Range * Sin(Elevation)
  z = -Range * Sin(Azimuth) * Cos(Elevation)

End Sub

'From Rectangular 3D coordinates to Spherical Coordinates (LH System)   (RADIANS)
Public Sub RecttoSph(ByRef Range As Single, ByRef Azimuth As Single, ByRef Elevation As Single, _
                     ByVal x As Single, ByVal y As Single, ByVal z As Single)

  Call RtoPQ(Range, Azimuth, x, -z)
  Call RtoPU(Range, Elevation, Range, y)

End Sub

'From Rectangular 2D coordinates to Polar Coordinates (RH system)      'Into correct Quadrant, Radians
Public Sub RtoPQ(ByRef R As Single, ByRef Angle As Single, ByVal x As Single, ByVal y As Single)

  R = Sqr(x * x + y * y)
  If x = 0 Then
    Angle = 0
    '    If y > 0 Then
    '      Angle = PI2         '90
    '    ElseIf y = 0 Then     '0
    '      Angle = 0
    '    Else
    '      y = PIPI - PI2      '270
    '    End If
  Else
    Angle = Atn(y / x)
    If Angle < 0 Then Angle = Angle + PI
    If y < 0 Then Angle = Angle + PI
  End If

End Sub

'From Rectangular 2D coordinates to Polar Coordinates (RH system)      'Unadorned Quadrant, Radians
Public Sub RtoPU(ByRef R As Single, ByRef Angle As Single, ByVal x As Single, ByVal y As Single)

  R = Sqr(x * x + y * y)
  If x = 0 Then
    Angle = 0
  Else
    Angle = Atn(y / x)
  End If

End Sub

'From Polar Coordinates to Rectangular 2D coordinates (LH system)      'Radians
Public Sub PtoR(ByRef x As Single, ByRef z As Single, ByVal R As Single, ByVal u As Single)

  x = R * Cos(u)
  z = -R * Sin(u)

End Sub

'This Maps a circle to the edges of the Texture square (0,0)-(1,1).  U is measured clockwise from (0.5,0)
'
'           (0,0)     ut      (1,0)
'                +-----|-----+
'                | / 7   0 \ |
'                |/6       1\|
'               -|     o     |- vt
'                |\5       2/|
'                | \ 4   3 / |
'                +-----|-----+
'           (0,1)             (1,1)

Public Sub MapUV(ByRef ut As Single, ByRef vt As Single, ByVal u As Single)

 Dim Octant As Integer

  u = u - PiPI * Int(u / PiPI)                                                     'u MOD 360
  Octant = Int(u / PI4)
  u = u - PI4 * Octant                                                             'Degrees into Octant

  Select Case Octant                                                               'which Octant
  Case 7: ut = 0.5 * (1 + Tan(u - PI4)): vt = 0                                  'TOP Edge
  Case 0: ut = 0.5 * (1 + Tan(u)):       vt = 0

  Case 1: ut = 1:                        vt = 0.5 * (1 + Tan(u - PI4))           'Right Edge
  Case 2: ut = 1:                        vt = 0.5 * (1 + Tan(u))

  Case 3: ut = 0.5 * (1 - Tan(u - PI4)): vt = 1
  Case 4: ut = 0.5 * (1 - Tan(u)):       vt = 1                                  'Bottom Edge

  Case 5: ut = 0:                        vt = 0.5 * (1 - Tan(u - PI4))
  Case 6: ut = 0:                        vt = 0.5 * (1 - Tan(u))                 'Left Edge

  End Select

End Sub

'An optimised routine to produce a fully saturated vector color
'u=0 gives Red, 60=Yellow, 120=Green, 180=Cyan, 240=Blue, 300=Magenta
'v=-90 saturation 1, 0=sat 0, 90 sat 1
Public Function UWtoRGB(ByVal u As Single, ByVal w As Single) As Long      'U,W in Radians

 Dim H As Single, q As Single, F As Single
 Dim t As Single, p As Single, s As Single
 Dim R As Single, g As Single, b As Single

  H = RtoD * u
  If H >= 360 Then H = H - 360
  If H < 0 Then H = H + 360
  H = H / 60
  F = H - Int(H)
  s = 0.8 * Abs(Sin(w)) + 0.2
  p = 1 - s
  q = 1 - s * F
  t = 1 - s * (1 - F)

  Select Case Fix(H)
  Case 0: R = 1: g = t: b = p
  Case 1: R = q: g = 1: b = p
  Case 2: R = p: g = 1: b = t
  Case 3: R = p: g = q: b = 1
  Case 4: R = t: g = p: b = 1
  Case 5: R = 1: g = p: b = q
  End Select

  UWtoRGB = (Int(255 * R) * 256& + Int(255 * g)) * 256 + Int(255 * b)

End Function

Public Function UtoRGB(ByVal u As Single) As Long      'U in Radians

  UtoRGB = UWtoRGB(u, PI2)

End Function

Public Sub ErrorMsgBox(ByRef Message As String, ByRef Err As ErrObject, ByVal Response As Long)

  Select Case Response
  Case vbCritical:
    MsgBox Message & vbCrLf & vbCrLf _
           & "Error was : " & Hex$(Err.Number) & "(" & Err.Number & ") " & Err.Description & vbCrLf & vbCrLf _
           & "This program will now END", vbCritical
  Case Else
    MsgBox Message & vbCrLf & vbCrLf _
           & "Error was : " & Hex$(Err.Number) & "(" & Err.Number & ") " & Err.Description & vbCrLf & vbCrLf _
           & "This program will continue"
  End Select

End Sub

'Fonts ----------------------------------------------------------------------------------------------------
Public Function CreateDXFont(ByRef NewFont As StdFont) As D3DXFont

 Dim OleFont As IFont
  
  Set OleFont = NewFont
  If Not OleFont Is Nothing Then
    Set CreateDXFont = D3DX.CreateFont(D3DD, OleFont.hFont)
  End If
End Function

Public Function CreateTextMesh(ByRef TextFont As StdFont, ByRef Text As String) As D3DXMesh
  Dim hDC As Long, hWndDC As Long
  Dim ohFnt As Long, OleFont As IFont
  Dim Mesh As D3DXMesh
      
  If Len(Text) = 0 Then Text = "???"
  hWndDC = GetDC(DXhWnd)
  If hWndDC Then
    hDC = CreateCompatibleDC(hWndDC)
    If hDC Then
      Set OleFont = TextFont
      If Not OleFont Is Nothing Then
        ohFnt = SelectObject(hDC, OleFont.hFont)
        On Local Error Resume Next
        Call D3DX.CreateText(D3DD, hDC, Text, 0.05, 0.2, CreateTextMesh, Nothing, ByVal 0)
        On Error GoTo 0
        Call SelectObject(hDC, ohFnt)
        Set OleFont = Nothing
      End If
      Call DeleteDC(hDC)
    End If
    Call ReleaseDC(DXhWnd, hWndDC)
  End If
End Function

':) Ulli's VB Code Formatter V2.13.5 (01-Feb-03 20:54:08) 58 + 236 = 294 Lines
