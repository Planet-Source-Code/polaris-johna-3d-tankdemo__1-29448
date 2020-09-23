VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7635
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   8880
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   8880
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'=================================================
'     (c) John Company Inc. 2000-2001
'
'
'    That is basic use of johnaDX7 engine
'    it shows how to initialize the Backbuffer and Zbuffer engine
'    it show how to caps user key input from keyboard
'    for explanation
'    write me at Johna.pop@caramail.com
'
'    if you like my engine write me and vote for me
'    if you want the engine source code email me
''
'
'
'===================================================
'
'
' This is the Version 1.01
'   it support Texture maping
'   -Multitexturing 2 level on my P2 333 with phonix Vanta
'   -Fog linar or exponentiel
'   -Terrain gererating
'   -Texture creation from JPEG,BMP,GIF files
'   -Directsound is support and Dsound3d
'   -Direct music is support
'   -light mapping in multitexturing FX tools method
'   -hight precision terrain colision detection
'   -Basic actor camera mov
'           -Walk
'           -turn arround
'           -jumping
'           -and many more
'
'   -Basic geometric object genarating as cube ect..
'   -and many others interesting useful stuff.....
'
'
'
'   -Now a Sound engine that handles with Dsound3d
'   -A particle engine system
'   - A beta vehicule AI engine for car,tank,ect..
'======================================================




Const PI = 3.145987421
Const RAY = 50



'===Objects declaration====


Dim DX7 As New johna_DX7
Dim KUB_1   'first cube object
Dim FLOOR() As D3DLVERTEX
Dim SKY As New cDome_sky  'the sky object
Dim FLOOR_BUMP As DirectDrawSurface7


Dim FLOOR_SURF As DirectDrawSurface7


Dim KUB_BUMP As DirectDrawSurface7

Dim MATH As New cMATH

'new particle engine
Dim SNOW_FALL As New cJohna_Particle
Dim SMOKE As New cJohna_Particle
Dim GEZER As New cJohna_Particle



  'change that constant for numTargets to be follown by the tank
Private Const MAX_TARGET = 50
Private Const MAX_Tank = 3 'use 1 or 2 to increase frame rate
Dim TANK As New XFileClass
Dim TANK_AI(MAX_Tank) As New cJohna_AI
Dim TANK_SMOKE(MAX_Tank) As New cJohna_Particle 'Tank smoke
Dim TANK_SOUND(MAX_Tank) As Integer 'tank bufferINDEX for the sound engine
'sound engine USE
Dim SOUND As New cJohna_Sound
Dim PLAYER_STEP As Integer
Dim PLAYER_SCREAM As Integer

Dim SOL(4) As D3DVERTEX

Dim FX As New Cjohna_Effects




Private Sub Form_Load()
Me.Refresh
Me.Show
'DX7.Initialize_DX_Windowed Me.hWnd  'initialize engine in windowed mode
'DX7.INiT_D3D Me.hWnd     'init the 3d engine
DX7.INIT_engineEX Me.hWnd
Call GAME_LOOP     'call the game loop
End Sub









'=========================================================
'  simple game loop
'  it finish when user toggle ESCAPE
'====================================================


Sub GAME_LOOP()
Dim RC As RECT
Dim zAngle As Single

'
   Dim V1 As D3DVECTOR
     Dim v2 As D3DVECTOR
      'for the land
     V1 = MATH.Vector(-3000, 0, -3000)
     v2 = MATH.Vector(3000, 0, 3000)
     
    DX7.DX_7.CreateD3DVertex V1.x, v2.Y, v2.Z, 0, 1, 0, 0, 0, SOL(0)
    DX7.DX_7.CreateD3DVertex v2.x, v2.Y, v2.Z, 0, 1, 0, 10, 0, SOL(1)
    DX7.DX_7.CreateD3DVertex V1.x, v2.Y, V1.Z, 0, 1, 0, 0, 10, SOL(2)
    DX7.DX_7.CreateD3DVertex v2.x, v2.Y, V1.Z, 0, 1, 0, 10, 10, SOL(3)
        







'======Create a simple Floor map

DX7.Make_PlaneSURF FLOOR(), DX7.johna_MakeVector(-500, 1, -500), DX7.johna_MakeVector(3500, 1, 3500), 254, 254, 255, 20, 20

'=======create the bump texture for multitexturing
Set FLOOR_BUMP = DX7.CreateTextureEX(App.Path + "\data\herbe_bump.bmp", 256, 256)
Set FLOOR_SURF = DX7.CreateTextureEX(App.Path + "\data\herbe1.bmp", 256, 256)




'=======Place the camera actor EYE the current orientation is the Degree=0
DX7.Camera_set_EYE DX7.johna_MakeVector(10, 10, 10)


'==============SKY object initialization

SKY.Init DX7, App.Path + "\data\ciel2.x", App.Path + "\data\sky03.jpg"

SKY.SKY_scale = 13


'================create a building------------------Position--------------Vector size------
KUB_1 = DX7.ADD_cubeEXTERN_face(DX7.johna_MakeVector(-200, 0, -200), DX7.johna_MakeVector(100, 250, 300), 250, 250, 250, App.Path + "\data\facade2.jpg", 1, 1)



'=======create the bump texture for multitexturing
Set KUB_BUMP = DX7.CreateTextureEX(App.Path + "\data\facade_bump.jpg", 256, 256)



'LIT = DX7.ADD_light(D3DLIGHT_SPOT, DX7.johna_MakeVector(10, 10, 10), DX7.D3Dcolor(0.1, 0.4, 0.5, 1), DX7.D3Dcolor(0.1, 0.4, 0.5, 1), 50, DX7.D3Dcolor(0.1, 0.4, 0.5, 1))


'Create snow particles
SNOW_FALL.Init DX7, MATH.Vector(0, 0, 0), App.Path + "\data\flare2.bmp", 100, 4, Johna_SNOW, Johna_STATIC_LOCATION, 0, 400, 400, 550
SMOKE.Init DX7, MATH.Vector(0, 0, 0), App.Path + "\data\smoke.bmp", 50, 150, Johna_SMOKE
GEZER.Init DX7, MATH.Vector(-2, 0, -150), App.Path + "\data\water_a.bmp", 400, 1 / 2, Johna_FONTAIN, , , , , , 5



'Init TANK
Dim I, J
Randomize Timer

For I = 0 To MAX_Tank
 TANK_AI(I).INIT_Ai MATH.Vector(Rnd * 120, 0, Rnd * 500), MATH.Vector(Rnd * 250, 0, Rnd * 250), 0.1, Rnd * 0.5, 4 * Rnd + 1.1, -1, , 1
 DX7.WriteLOG "Add tank" + Str(I) + " AI", Me.hWnd
Next I



For J = 0 To MAX_Tank
For I = 1 To MAX_TARGET
  TANK_AI(J).AI_Add_Target MATH.Vector(Rnd * 1500, 0, Rnd * 1500)
  
Next I

DX7.WriteLOG "Add tank Targets for tank #" + Str(J), Me.hWnd
Next J
TANK.Load DX7, App.Path + "\data\tank.x"
Dim VC As D3DVECTOR
Dim VC2 As D3DVECTOR
DX7.WriteLOG "Add tank smoke particle", Me.hWnd
'init tank particle
For I = 0 To MAX_Tank
 TANK_SMOKE(I).Init DX7, MATH.Vector(0, 0, 0), App.Path + "\data\smoke.bmp", 20, 4, Johna_SMOKE
 
Next


'init sound
DX7.WriteLOG "Tank sound", Me.hWnd

SOUND.INIT_SoundENGINE Me.hWnd
PLAYER_STEP = SOUND.Load_Wav(App.Path + "\data\SnowStep1.wav", True)
PLAYER_SCREAM = SOUND.Load_Wav(App.Path + "\data\screamshort.wav", True)

For I = 0 To MAX_Tank

  TANK_SOUND(I) = SOUND.Load_Wav(App.Path + "\data\tankidle.wav", True)
  
Next I


For I = 0 To MAX_Tank

  SOUND.Play_BUF TANK_SOUND(I), True
Next I

Do


  zAngle = zAngle + 0.8
  If zAngle > 360 Then zAngle = 0

  DoEvents
  If DX7.GetKEY(DIK_ESCAPE) Then GoTo END_it
  'checkeys
  Call Me.KEY_check
  
  
  
  
 SOUND.Set_ListenerPosition DX7.GET_CameraEYE
 SOUND.Set_BufferPosition PLAYER_STEP, DX7.GET_CameraEYE



  'Render 3d
  DX7.D3D_DEV.BeginScene
  DX7.Clear_3D
  
  
  
  
  
  'tank
  For I = 0 To MAX_Tank
  Dim Y As Single
  Dim Z As Single
    TANK_AI(I).Update_AI
    VC2 = TANK_AI(I).Get_location
    
    'update each tank position and rotation
    TANK.Vrotation = MATH.Vector(Z, Round(TANK_AI(I).GetAngle_DEG), 0)
    TANK.Vscal = MATH.Vector(5, 5, 5)
    TANK.vPosition = VC2
    'render a single tank
    TANK.Render DX7
    'render tank Smoke
    TANK_SMOKE(I).SetPosition VC2
    TANK_SMOKE(I).Render DX7
    SOUND.Set_BufferPosition TANK_SOUND(I), VC2
    
    'Test colision
    If TANK.IScollision(DX7.GET_CameraEYE) Then
       SOUND.Set_BufferPosition PLAYER_SCREAM, DX7.GET_CameraEYE
       SOUND.Play_BUF PLAYER_SCREAM
    End If
  Next I
  DX7.CLearMATRIX
  
  DX7.MULTITEXTURE_FX_Dark_Mapping 1, FLOOR_BUMP
  DX7.D3D_DEV.SetTexture 0, FLOOR_SURF
  DX7.D3D_DEV.DrawPrimitive D3DPT_TRIANGLESTRIP, D3DFVF_VERTEX, SOL(0), 4, D3DDP_DEFAULT
  DX7.Disable_MULTITEXTURE
  
  
  DX7.MULTITEXTURE_FX_Dark_Mapping 1, KUB_BUMP
  DX7.Render_EXTERNcube KUB_1, 1
  DX7.Disable_MULTITEXTURE
  
  
   'render Sky
   SKY.Render DX7
  
   'render snow particle
   SNOW_FALL.Render DX7
  
  GEZER.SetPosition MATH.Vector(50 + RAY * Cos(zAngle * 1 / 2 * PI / 180), 0, -50 + RAY * Sin(zAngle * 1 / 2 * PI / 180))
  GEZER.Render DX7
  
  
  
 
  
  'animate smoke effect
  SMOKE.SetPosition MATH.Vector(0 + RAY * Cos(zAngle * PI / 180), 10, 0 + RAY * Sin(zAngle * PI / 180))
  SMOKE.Render DX7
  
  
  
  
  
  DX7.D3D_DEV.EndScene
  'end rendering


  DX7.BAK.DrawText 10, 40, "FramePer seconde=" + Str(DX7.FramesPerSec), 0
  DX7.BAK.DrawText 10, 60, "PRESS SPACE for reset camera", 0
  
  
  
  DX7.FLIPP Me.hWnd
Loop


END_it:

DX7.FreeDX Me.hWnd
End



End Sub





Sub KEY_check()


If DX7.GetKEY(DIK_UP) Then
  DX7.Camera_Move_Foward 1
 
  SOUND.Play_BUF PLAYER_STEP
 
End If
  
If DX7.GetKEY(DIK_RCONTROL) Then
     DX7.Camera_Move_Foward 8
     SOUND.Play_BUF PLAYER_STEP
    
End If

If DX7.GetKEY(DIK_DOWN) Then
  
     SOUND.Play_BUF PLAYER_STEP
     DX7.Camera_Move_Backward 1
End If

If DX7.GetKEY(DIK_LEFT) Then _
  DX7.Camera_Move_Left 0.0005

If DX7.GetKEY(DIK_RIGHT) Then _
  DX7.Camera_Move_Right 0.0005


If DX7.GetKEY(DIK_NUMPAD8) Then _
  DX7.Camera_Move_UP 0.0005

If DX7.GetKEY(DIK_NUMPAD2) Then _
  DX7.Camera_Move_DOWN 0.0005



If DX7.GetKEY(DIK_NUMPAD4) Then _
  DX7.CAM_step_LEFT 1

If DX7.GetKEY(DIK_NUMPAD6) Then _
  DX7.CAM_step_RIGHT 1


If DX7.GetKEY(DIK_SPACE) Then _
  DX7.Camera_set_EYE DX7.johna_MakeVector(10, 10, 10)

If DX7.GetKEY(DIK_ADD) Then DX7.Camera_Elevator_UP 1
If DX7.GetKEY(DIK_SUBTRACT) Then DX7.Camera_Elevator_DOWN 1




Dim VC As D3DVECTOR
Dim TempVC As D3DVECTOR

VC = DX7.GET_CameraEYE

'VC.y = Land.Get_Altitude_EX(VC) + 10

DX7.Camera_set_EYE VC
SOUND.Set_ListenerPosition VC
End Sub

