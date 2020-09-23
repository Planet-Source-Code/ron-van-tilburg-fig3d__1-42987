VERSION 5.00
Begin VB.Form MainForm 
   Caption         =   "Fig3D ©2003 RVT - DirectX Demos -  Hit ESC to Finish"
   ClientHeight    =   6795
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9480
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   12
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   ScaleHeight     =   453
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   632
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Fig3D - a Demonstration of the capabiltites of the rvtDX.dll  D3D graphics engine
'©2003 Ron van Tilburg - rivit@f1.net.au
'Freeware for Educational Purposes, For commercial interests contact author please, I retain copyright.

'======================================================================================================
'Comment Out that which you arent showing
'I havent figured out how to get back to a normal window from Full Screen without hanging yet!

'TODO:
' Prescaling, Framerate control, Improved Documentation, More Figures, XFiles in and Out
' Improve Control of Texture Rendering, Other effects incl. Fog, Tuning, Making My Own Meshes
' etc. etc. Help, Suggestions and Example Scenes Welcome there's tons more to learn

'VIEW CONTROL:
'When the Axes are shown you can usually affect the view in the following ways with the following keys:

'    Escape:      exit (this part of) the program
'    PageUp:      viewpoint up
'    PageDown:    viewpoint down
'    Home:        viewpoint clockwise
'    End:         viewpoint anticlockwise
'    Delete:      zoom out
'    Insert:      zoom in
'    ScrollLock   reset viewpoint to centre

'PERFORMANCE:
'This was built and tested on a 1999 Dell XPS600T Win98 machine (P3 600Mhz, 384MB), _
'and a nVidia AGP 2x, 32MB TNT2 GFx card (Detonator 4109 drivers), with DirectX 8.1a, _
'Max FPS appears to be 120 (done while rendering practically nothing)

'For these specs I get the following Framerates - your newer, faster, machine will probably kill it
'However the animations may look a bit strange (frantic) - I havent attempted to control frame rates yet
        
'  Resolution         FPS    Figures (anim)            Saturn (anim)        Earth (anim)
'                        InVBWin FullScreen       InVBWin FullScreen    InVBWin FullScreen
'   640* 480*32            29.7    36.8             65.7    118.1*       59.6    78.1      *rendering looks odd
'   800* 600*32            21.8    22.9             63.1    105.6        58.5    75.6
'  1024* 768*32            19.8    20.4             47.3     72.9        49.9    70.5
'  1280*1024*32            12.7    14               29.7     43.9        33.1    52.1      'Target >=24
'  1600*1200*32             8.2    9.6              21.6     31.1        25      38.2

'It is probable these could be tuned to be better ...   RVT 2 Feb 2003
'======================================================================================================

Public rvtDX As cDX8              'THE ONLY GLOBAL WE NEED FOR DX8

Private Const PlayInVBWindow As Boolean = True      'Change This to False to Go FullScreen

Private Sub Form_Load()

  Call Show
  Set rvtDX = New cDX8
  If rvtDX.InitDX8(Me.hWnd) Then
    If PlayInVBWindow Then MaximizeWindow   'Maximise It (not relly necessary but does show max resolution off)
    
    If rvtDX.GetRenderDevice(InVBWindow:=PlayInVBWindow, Depth:=-1) Then 'Depth=-1 forces use of Default Display in Fullscreen
      
      'If you want some other background sound replace this file with one of your own (has to be a .wav)
      'Call rvtDX.PlaySoundFromFile(App.Path & "\textures\Loop.wav", LoopIt:=True)
      
      'COMMENT THE ONE(s) YOU DONT WISH TO SEE
      ShowFigures
      ShowSaturnFlyby
      ShowEarth
    End If
  End If
  Call Unload(Me)

End Sub

Private Sub MaximizeWindow()

  MainForm.Left = 0
  MainForm.Top = 0
  MainForm.Width = Screen.Width
  MainForm.Height = Screen.Height
  MainForm.WindowState = vbMaximized

End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  
  If UnloadMode = vbFormControlMenu Then Cancel = True

End Sub

Private Sub Form_Unload(Cancel As Integer)

  Set rvtDX = Nothing                           'blow away everything to do with DirectX

End Sub

'=====================================================================================================
'=====================================================================================================

':) Ulli's VB Code Formatter V2.13.5 (31-Jan-03 12:20:41) 3 + 33 = 36 Lines
