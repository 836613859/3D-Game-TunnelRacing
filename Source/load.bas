Attribute VB_Name = "f¼ÓÔØ"
'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª°ëÍ¸Ã÷´°ÌåAPI¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function SetCursorPos Lib "user32" (ByVal X As Long, ByVal Y As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Const WS_EX_LAYERED = &H80000
Public Const GWL_EXSTYLE = (-20)
Public Const LWA_ALPHA = &H2
Public Const LWA_COLORKEY = &H1 '²»ÏÔÊ¾Í¸Ã÷É« - -ÍÚ¿ÕÁË°¡Ì«Ë¬ÁË
'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª





Public TV As TVEngine 'ÉùÃ÷TV3DÒýÇæ¶ÔÏó
Public InputE As TVInputEngine 'ÉùÃ÷ÊäÈë£¨¼üÅÌ¡¢Êó±ê£©¶ÔÏó

Public mesh As TVMesh
Public sd(10) As TVMesh
Public box(110) As TVMesh
Public planet(10) As TVMesh
Public razer(10) As TVMesh
Public curtain As TVMesh
Public Dead_Razer As TVMesh
Public body As TVMesh
Public PrtSurface As TVMesh

Public scene As TVScene
Public cam As TVCamera
Public tex As TVTextureFactory
Public Atmo As TVAtmosphere
Public effect As TVGraphicEffect
Public lightE As New TVLightEngine
Public LightD(11) As D3DLIGHT8

Public soundE As New TVSoundEngine
Public sound As New TVSoundMP3
Public bgm As New TVSoundMP3
Public GMusic As New TVSoundMP3
Public HpSound As New TVSoundMP3
Public RazerSound As New TVSoundMP3

Public HasPlayedMovie As Boolean

'¡ª¡ª¡ª¡ª¡ª¡ª±äÁ¿¶¨Òå¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª
Public mNum As Integer
Public GPath As String
Public T As Long 'ÎÈ¶¨µÄÊ±¼äÁ÷ÊÅ
Public fps As Long
Public lastT As Long 'ÉÏÒ»´ÎµÄÊ±¼äÁ÷ÊÅ
Public sdNum As Long

Public SpeedUpNum As Long '¼ÓËÙµÀ¾ßÊý
Public SpeedDownNum As Long
Public HPNum As Long
Public ProtectNum As Long
Public MoneyOwn As Variant
Public HighScore As Variant

Public fµÝÔöÖµ As Single 'ºá£¬Êú£¬ÍÖÔ²ÔË¶¯µÄÕÏ°­ÔË¶¯ÓÃµ½µÄ³£Êý£¬²»¶ÏµÝÔöµÄ
Public CamX As Long
Public CamZ As Long
Public CamY As Long
Public StaySame As Single '
Public Vcam As Single 'cameraÒÆ¶¯ËÙ¶È
Public Scam As Single 'ËíµÀÖÐÒÆ¶¯Â·³Ì
Public STcam As Single 'cameraÂ·³Ì
Public RVel As Single 'cameraË®Æ½¼ÓËÙ¶È
Public UPVel As Single 'camera´¹Ö±¼ÓËÙ¶È
Public Rs As Long
Public LightID As Integer
Public HP As Single
Public Died As Boolean 'ËÀÁËÃ»
Public HasFlash As Boolean 'ËÀÁËÖ®ºóÉÁÁËÃ»..
Public isBeingProtecting  As Boolean '·À»¤ÕÖ¿ªÁËÃ»
Public Tools(4) As Integer ' µÀ¾ß
Public razer_shoot As Boolean 'Å£X¼¤¹â¿ÉÒÔÉäÃ»
Public razerX As Single
Public razerY As Single
Public razerZ As Single
Public sf_alpha As Single '·À»¤ÕÖALPHA
Public HasUnloadForm1 As Long

Public al As Long 'ÆÀ·ÖÆÁµÄTop
Public GMusicPlaying As Boolean 'ÓÎÏ·±³¾°ÒôÀÖÓÐÃ»ÔÚ²¥·Å
Public SleepTimeLeft As Single 'lockFPS

Sub ¼ÓÔØ()


'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª±äÁ¿¡ª¡ª¡ª¡ª¡ª¡ª
Vcam = 8
Scam = 0
STcam = 0
RVel = 0
UPVel = 0
HP = 100
mNum = 1
Ttotal = 0
StaySame = 0
T = 0
TL = 0
sf_alpha = 0.2
al = -12000
razerX = 40000
razer_shoot = False
isBeingProtecting = False
Died = False
 HasFlash = False
  HasWrittenDownZY = False

GPath = App.Path & "\data\"
'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª

'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ªset¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª
Set TV = New TVEngine
Set scene = New TVScene
Set tex = New TVTextureFactory
Set InputE = New TVInputEngine
Set cam = New TVCamera
Set Atmo = New TVAtmosphere
Set effect = New TVGraphicEffect
Set body = scene.CreateMeshBuilder
Set GMusic = New TVSoundMP3

'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª

'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ªÒýÇæÉèÖÃ¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª
TV.Init3DWindowedMode Form1.Picture1.hwnd
'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª
'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª²ÄÖÊ¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª
tex.LoadTexture GPath & "wall.jpg", "wall", 200, 200
tex.LoadTexture GPath & "dirt.bmp", "dirt"
tex.LoadTexture GPath & "iron2.jpg", "iron"
tex.LoadTexture GPath & "kdp.jpg", "kdp", 20, 17
tex.LoadTexture GPath & "front.jpg", "front"
tex.LoadTexture GPath & "left.jpg", "left"
tex.LoadTexture GPath & "right.jpg", "right"
tex.LoadTexture GPath & "back.jpg", "back"
tex.LoadTexture GPath & "top.jpg", "top"
tex.LoadTexture GPath & "down.jpg", "down"
'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª-¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª

'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ªÌì¿ÕºÐ¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª
Atmo.SkyBox_SetTexture GetTex("front"), GetTex("back"), GetTex("left"), GetTex("right"), GetTex("top"), GetTex("down")
Atmo.SkyBox_SetDistance 10000
  Atmo.SkyBox_Enable True
'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª


'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ªËíµÀ¼ÓÔØ¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª
For sdNum = 1 To 10
Set sd(sdNum) = scene.CreateMeshBuilder
sd(sdNum).Load3DSMesh GPath & "sd.3ds"
sd(sdNum).SetPosition 9600 - 9600 * sdNum, 0, 0
sd(sdNum).ScaleMesh 2, 1, 1
sd(sdNum).SetTexture GetTex("wall")
Next sdNum
'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª

'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ªÐÇÇò¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª
For plnum = 1 To 5
Set planet(plnum) = scene.CreateMeshBuilder
planet(plnum).Load3DSMesh GPath & "c.3ds"
planet(plnum).SetTexture GetTex("dirt")
Next
planet(1).SetPosition -5000, 200, -2000
planet(2).SetPosition -46000, 200, -5000
planet(3).SetPosition -10000, 200, 2000
planet(4).SetPosition -30000, 200, 4500
planet(5).SetPosition -36000, 200, 5000

'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª
'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª·À»¤ÕÖ¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª
Set PrtSurface = scene.CreateMeshBuilder
PrtSurface.Load3DSMesh GPath & "c.3ds"
PrtSurface.SetColor RGBA256(0, 0, 100, 0.2)
PrtSurface.SetPosition 3000, 0, 0
PrtSurface.ScaleMesh 0.05, 0.05, 0.05
'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª


'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ªºÐ×Ó¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª
For boxnum = 1 To 100
Set box(boxnum) = scene.CreateMeshBuilder
box(boxnum).Load3DSMesh GPath & "box.3ds"
Randomize

box(boxnum).SetPosition Int(-Rnd(1) * 77000) - 2000, Int(-Rnd(1) * 250) + 150, Int(-Rnd(1) * 150) + 70
box(boxnum).SetTexture GetTex("iron")
box(boxnum).ScaleMesh Rnd(1) * 3 + 0.3, Rnd(1) * 3 + 0.3, Rnd(1) * 3 + 0.3
Next
For boxnum = 101 To 110
Set box(boxnum) = scene.CreateMeshBuilder
box(boxnum).Load3DSMesh GPath & "box.3ds"
box(boxnum).SetTexture GetTex("iron")
box(boxnum).SetPosition 10000, 0, 0
Next
'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ªµÆ¹â¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª

    'Dim LightD(10) As D3DLIGHT8
   
    For lightnum = 1 To 10
    LightD(lightnum).Type = D3DLIGHT_POINT
    LightD(lightnum).Position = Vector(-30000 - 5000 * lightnum, 70, 0)
    LightD(lightnum).Ambient = DXColor(0, 300, 300, 5)
   LightD(lightnum).diffuse = DXColor(0, 300, 300, 5)
    LightD(lightnum).Range = 7000
    LightD(lightnum).Attenuation0 = 0.0008
    LightD(lightnum).Attenuation1 = 0.0008
  LightD(lightnum).Attenuation2 = 0.0008
   lightE.CreateLight LightD(lightnum)
Next
LightD(11).Type = D3DLIGHT_POINT
LightD(11).Position = Vector(0, 70, 0)
    LightD(11).Ambient = DXColor(300, 300, 300, 5)
   LightD(11).diffuse = DXColor(300, 300, 300, 5)
    LightD(11).Range = 46000
    LightD(11).Attenuation0 = 0
    LightD(11).Attenuation1 = 0
  LightD(11).Attenuation2 = 0
   lightE.CreateLight LightD(11)
'lightE.CreateQuickPointLight Vector(0, 70, 0), 300, 300, 300, 46000

'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª

'¡ª¡ª¡ª¡ª-¡ª¡ª¡ª¡ª¡ª¡ª¿ÓµùÆÁ¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª
Set curtain = scene.CreateMeshBuilder
curtain.SetBillboardType TV_BILLBOARD_NOROTATION
curtain.Load3DSMesh GPath & "box.3ds"
curtain.SetColor RGBA(0.5, 0.5, 0.5, 0.7)
curtain.ScaleMesh 0.01, 50, 5
curtain.SetPosition 10000, 0, 0
'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª ¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª

'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¼¤¹âÊø¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª
For rznum = 1 To 10
Set razer(rznum) = scene.CreateMeshBuilder
razer(rznum).Load3DSMesh GPath & "box.3ds"
razer(rznum).ScaleMesh 0.4, 0.4, 5
razer(rznum).SetColor RGBA(300, 0, 0, 0.6)
razer(rznum).SetPosition 1000, 0, 0
Next

'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª

'¡ª¡ª¡ª¡ª¡ª¡ªÅ£X¼¤¹â£¬10ÃëÀ´Ò»·¢¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª
Set Dead_Razer = scene.CreateMeshBuilder
Dead_Razer.Load3DSMesh GPath & "box.3ds"
Dead_Razer.ScaleMesh 1000, 0.5, 0.5
Dead_Razer.SetColor RGBA(0, 300, 300, 0.9)
Dead_Razer.SetPosition 0, 1000, 0

'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª


'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ªcamera³õÉèÖÃ¡ª¡ª¡ª¡ª¡ª¡ª
CamX = 300
CamY = 0
CamZ = 0

cam.SetCamera 300, 0, 0, 0, 0, 0
cam.SetViewFrustum 70, 10000
body.CreateBox 30, 60, 60
body.SetPosition 400, 50, 50
'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª
'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ªÓÎÏ·ÒôÀÖ  ÒôÐ§¡ª¡ª¡ª¡ª¡ª¡ª
GMusic.Load GPath & "bgm/" & mNum & ".mp3"
GMusic.Play
GMusic.Volume = 0
 GMusicPlaying = True
 
HpSound.Load GPath & "snd/HP.mp3"
RazerSound.Load GPath & "snd/Razer.mp3"

 
'¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª¡ª



effect.FadeIn 5000

End Sub
