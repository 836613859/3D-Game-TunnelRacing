Attribute VB_Name = "游戏中"
Private razerX As Single
Private razerY As Single
Private razerZ As Single
Private HasWrittenDownZY As Boolean

Sub f开始()
If T < 200 Then CamX = CamX - Vcam * (T / 10)
Scam = Scam + Vcam * (T / 10)
STcam = STcam + Vcam * (T / 10)
If T > 200 Then Scam = Scam - Vcam * (T / 10)  '- 5
Vcam = Vcam + 0.001 * TV.TimeElapsed / 10
f递增值 = f递增值 + 0.001 * T / 10
Form1.Label1.Caption = "速度：" & Left(Vcam, 5) & " m/s"
Form1.Label2.Caption = "生命：" & Int(HP)
Form1.Label3.Caption = "FPS:" & TV.GetFPS
End Sub
Sub f按键检测()
    
    '————————判断按键———————`———
    T = 1000 / (TV.GetFPS + 1)
    fps = TV.GetFPS + 1
    InputE.EnableEvents True

    If InputE.IsKeyPressed(TV_KEY_DOWN) = True Then
    UPVel = UPVel - 0.7 * (T / 10)
    StaySame = 0
    End If
    
    If InputE.IsKeyPressed(TV_KEY_UP) = True Then
    UPVel = UPVel + 0.7 * (T / 10)
    StaySame = 0
    End If
    
   If InputE.IsKeyPressed(TV_KEY_LEFT) = True Then
    RVel = RVel - 0.7 * (T / 10)
    StaySame = 0
    End If
    
    If InputE.IsKeyPressed(TV_KEY_RIGHT) = True Then
    RVel = RVel + 0.7 * (T / 10)
    StaySame = 0
    End If
    
    If InputE.IsKeyPressed(TV_KEY_PAGEDOWN) = True Then
    mNum = mNum + 1
    If mNum = 7 Then mNum = 1
    Set GMusic = Nothing
    Set GMusic = New TVSoundMP3
   GMusic.Load GPath & "bgm/" & mNum & ".mp3"
   GMusic.Play
    End If
    
    If InputE.IsKeyPressed(TV_KEY_PAGEUP) = True Then
    mNum = mNum - 1
    If mNum = 0 Then mNum = 6
    Set GMusic = Nothing
    Set GMusic = New TVSoundMP3
   GMusic.Load GPath & "bgm/" & mNum & ".mp3"
   GMusic.Play
    End If

        '——————————————道具————————————————
    If InputE.IsKeyPressed(TV_KEY_A) And Died = False And HPNum > 0 And HP < 100 Then
    HPNum = HPNum - 1
    Form1.L_HP.Caption = HPNum
    HpSound.Play
   HP = 100
    End If
    
     If InputE.IsKeyPressed(TV_KEY_S) And Vcam >= 1.1 And Died = False And SpeedDownNum > 0 Then
     SpeedDownNum = SpeedDownNum - 1
     Form1.L_SD.Caption = SpeedDownNum
    Vcam = Vcam - 0.3
    End If

    If InputE.IsKeyPressed(TV_KEY_W) And SpeedUpNum > 0 And Died = False Then
    SpeedUpNum = SpeedUpNum - 1
    Form1.L_Sup.Caption = SpeedUpNum
    Vcam = Vcam + 0.5
    End If
    
    If InputE.IsKeyPressed(TV_KEY_D) And ProtectNum > 0 And Died = False And isBeingProtecting = False Then
    ProtectNum = ProtectNum - 1
    Form1.L_P.Caption = ProtectNum
    isBeingProtecting = True
    Form1.Timer4.Enabled = True
    End If
    '————————————————————————————————————
    
    If InputE.IsKeyPressed(TV_KEY_DOWN) = False And InputE.IsKeyPressed(TV_KEY_UP) = False Then
        If UPVel > 0 Then UPVel = UPVel - 3 * (T / 10) / fps
        If UPVel < 0 Then UPVel = UPVel + 3 * (T / 10) / fps
        If UPVel < 2 * (T / 10) / fps And UPVel > -2 * (T / 10) / fps Then UPVel = 0
    End If
    
    If InputE.IsKeyPressed(TV_KEY_LEFT) = False And InputE.IsKeyPressed(TV_KEY_RIGHT) = False Then
        If RVel > 0 Then RVel = RVel - 3 * (T / 10) / fps
        If RVel < 0 Then RVel = RVel + 3 * (T / 10) / fps
        If RVel < 2 * (T / 10) / fps And RVel > -2 * (T / 10) / fps Then RVel = 0
    End If


        Select Case UPVel '上下左右的速度限定
    Case Is < -2 * (T / 10)
    UPVel = -2 * (T / 10)
    Case Is > 2 * (T / 10)
    UPVel = 2 * (T / 10)
    End Select
    
        Select Case RVel
    Case Is < -2 * (T / 10)
    RVel = -2 * (T / 10)
    Case Is > 2 * (T / 10)
   RVel = 2 * (T / 10)
    End Select
    
 '  Open "d:\1.txt" For Append As #1
 '  Print #1, "UPVEL " & UPVel & "     " & "RVEL" & RVel & "T:" & T
 ' Close #1
    
    
    CamY = CamY + UPVel
    CamZ = CamZ + RVel
    '————————按键结束—————————
End Sub
Sub f边界检测()

If CamZ <= -110 Then CamZ = -110
If CamZ >= 110 Then CamZ = 110
If CamY <= -70 Then CamY = -70
If CamY >= 70 Then CamY = 70

    cam.SetPosition CamX, CamY, CamZ
    body.SetPosition CamX + 31, CamY, CamZ
      
  If razer_shoot = False Then
  razerX = CamX + 40000
  HasWrittenDownZY = False
  Dead_Razer.SetPosition razerX, 0, 0
  End If
  
  If isBeingProtecting = True Then
  sf_alpha = sf_alpha - TV.TimeElapsed * 0.008
  PrtSurface.SetPosition CamX + 10, CamY + 10, CamZ + 10
  PrtSurface.SetColor RGBA256(0, 0, 100, sf_alpha)
  End If
End Sub

Sub f回归起点()
If Scam > 85000 Then
 CamX = 0
  Scam = 0
 cam.SetPosition CamX, CamY, CamZ

End If

End Sub
Sub f渲染()

    TV.Clear
    
  scene.RenderAllMeshes
    Atmo.Skybox_Render
    TV.RenderToScreen
    
End Sub
Sub f障碍重置()
If Scam > 84500 And Scam < 85000 Then '到隧道末

  Randomize
  For boxnum = 1 To 100
  box(boxnum).SetPosition Int(-Rnd(1) * 77000) - 2000, Int(-Rnd(1) * 250) + 150, Int(-Rnd(1) * 150) + 70
  Next boxnum

  For lightnum = 1 To 10
  LightD(lightnum).Ambient = DXColor(Rnd(1) * 300 + 50, Rnd(1) * 300 + 50, Rnd(1) * 300 + 50, 5)
  LightD(lightnum).diffuse = DXColor(Rnd(1) * 300 + 50, Rnd(1) * 300 + 50, Rnd(1) * 300 + 50, 5)
  lightE.CreateLight LightD(lightnum)
  Next
  lightE.CreateQuickPointLight Vector(0, 70, 0), 300, 300, 300, 45000

     Randomize
     curtain.SetPosition Int(-Rnd(1) * 50000 + 10000), 0, 0

 'If Vcam > 12 Then
     For rznum = 1 To 10
     razer(rznum).SetPosition -20000 - rznum * 800, -Rnd(1) * 160 + 80, 0
     Next rznum
 'End If

 effect.FadeIn 3000
End If

If Vcam > 10 Then
For boxnum = 10 To 70 Step 6
box(boxnum).SetPosition box(boxnum).GetPosition.X, box(boxnum).GetPosition.Y, -30 + 150 * Sin(f递增值 * 30)
Next
End If

If Vcam > 13 Then
For boxnum = 30 To 90 Step 6
box(boxnum).SetPosition box(boxnum).GetPosition.X, -50 + 100 * Sin(f递增值 * 30), box(boxnum).GetPosition.z
Next
Form1.Timer3.Enabled = True
End If

If Vcam > 16 Then
For boxnum = 20 To 80 Step 5
box(boxnum).SetPosition box(boxnum).GetPosition.X, -30 + 80 * Sin(f递增值 * 30 - 130), -40 + 80 * Sin(f递增值 * 30)
Next
End If

If Vcam > 18 Then
Form1.Timer3.Interval = 10000
End If

If razer_shoot = True Then '射出牛X激光
razerX = razerX - 300 * T / 10
  If HasWrittenDownZY = False Then
  razerY = CamY
  razerZ = CamZ
  HasWrittenDownZY = True
  End If
  Dead_Razer.SetPosition razerX, razerY, razerZ
  If razerX < CamX - 40000 Then
  razerX = CamX + 40000
  razer_shoot = False
  HasWrittenDownZY = False
  Dead_Razer.SetPosition razerX, 0, 0
  GMusic.Volume = 0
  End If
  
End If

If UPVel = 0 And RVel = 0 Then StaySame = StaySame + 0.1 * TV.TimeElapsed
If StaySame > 300 And CamX < box(110).GetPosition.X Then
For boxnum = 101 To 110
box(boxnum).SetPosition CamX - 1000, -70 + 15 * (boxnum - 100), CamZ
Next
StaySame = 0
End If
End Sub


Sub f碰撞测试()

If isBeingProtecting = True Then GoTo ex:

For clnum = 1 To 110
If box(clnum).TestBoxCollideWith(body) = True Then
HP = HP - Vcam ^ 2 / 30
Vcam = Vcam - 0.05
effect.Flash 0.7, 0, 0, 500
End If
Next

For clnum2 = 1 To 10
If razer(clnum2).TestBoxCollideWith(body) = True Then
HP = HP - 10
effect.Flash 0.7, 0, 0, 500
End If
Next

If Vcam < 1 Then Vcam = 1

If Dead_Razer.TestBoxCollideWith(body) = True Then
HP = HP - 5
effect.Flash 0.7, 0, 0, 500
End If

If HP <= 0 Then
HP = 0
Died = True
HasFlash = False
End If

ex:

End Sub


