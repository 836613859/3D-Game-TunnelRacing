VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "隧道狂飙"
   ClientHeight    =   10485
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   15855
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   14.25
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form1.frx":324A
   ScaleHeight     =   10485
   ScaleWidth      =   15855
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   15000
      Left            =   11160
      Top             =   1080
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   30
      Left            =   11625
      Top             =   525
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000007&
      Caption         =   "选 项"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   3765
      Left            =   9525
      TabIndex        =   4
      Top             =   3300
      Visible         =   0   'False
      Width           =   3690
      Begin VB.CommandButton Command4 
         BackColor       =   &H00FFFFFF&
         Caption         =   "返回主菜单"
         Height          =   765
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   2775
         Width           =   3240
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H80000014&
         Caption         =   "开始/暂停音乐"
         Height          =   765
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   1920
         Width           =   3240
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H80000014&
         Caption         =   "重新开始"
         Height          =   690
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   1200
         Width           =   3240
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "继续"
         Height          =   705
         Left            =   225
         Style           =   1  'Graphical
         TabIndex        =   5
         Top             =   450
         Width           =   3240
      End
   End
   Begin VB.PictureBox 游戏结束评分栏 
      BackColor       =   &H0080FFFF&
      Height          =   8565
      Left            =   6750
      ScaleHeight     =   8505
      ScaleWidth      =   8430
      TabIndex        =   13
      Top             =   2550
      Visible         =   0   'False
      Width           =   8490
      Begin VB.CommandButton 返回按钮 
         BackColor       =   &H0080FF80&
         Caption         =   "返回主菜单"
         Height          =   915
         Left            =   5025
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   6750
         Width           =   2415
      End
      Begin VB.CommandButton 重来按钮 
         BackColor       =   &H0080FF80&
         Caption         =   "重新开始"
         Height          =   840
         Left            =   975
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   6750
         Width           =   2340
      End
      Begin VB.Label L_HighScore 
         BackStyle       =   0  'Transparent
         Caption         =   "路程 最高记录："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   465
         Left            =   600
         TabIndex        =   24
         Top             =   5520
         Width           =   6390
      End
      Begin VB.Image Image1 
         Height          =   450
         Left            =   3900
         Picture         =   "Form1.frx":39DC
         Top             =   4800
         Width           =   840
      End
      Begin VB.Label 奖励金币 
         BackStyle       =   0  'Transparent
         Caption         =   "奖励XO币（有速度加成）："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   1065
         Left            =   600
         TabIndex        =   17
         Top             =   4350
         Width           =   6690
      End
      Begin VB.Label 总路程 
         BackStyle       =   0  'Transparent
         Caption         =   "总路程："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   465
         Left            =   600
         TabIndex        =   15
         Top             =   2025
         Width           =   5715
      End
      Begin VB.Label 标题 
         BackStyle       =   0  'Transparent
         Caption         =   "游 戏 结 束"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   42
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1065
         Left            =   1800
         TabIndex        =   14
         Top             =   450
         Width           =   5340
      End
      Begin VB.Label 末速度 
         BackStyle       =   0  'Transparent
         Caption         =   "最后的速度："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H008080FF&
         Height          =   465
         Left            =   600
         TabIndex        =   16
         Top             =   3225
         Width           =   6390
      End
   End
   Begin VB.PictureBox Picture5 
      Height          =   915
      Left            =   1050
      Picture         =   "Form1.frx":8DC5
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   12
      Top             =   6300
      Width           =   915
   End
   Begin VB.PictureBox Picture2 
      Height          =   915
      Left            =   1050
      Picture         =   "Form1.frx":A645
      ScaleHeight     =   855
      ScaleWidth      =   780
      TabIndex        =   11
      Top             =   4500
      Width           =   840
   End
   Begin VB.PictureBox Picture3 
      Height          =   915
      Left            =   75
      Picture         =   "Form1.frx":B6D1
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   10
      Top             =   6300
      Width           =   915
   End
   Begin VB.PictureBox Picture4 
      Height          =   915
      Left            =   75
      Picture         =   "Form1.frx":C50F
      ScaleHeight     =   855
      ScaleWidth      =   855
      TabIndex        =   9
      Top             =   4500
      Width           =   915
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   11100
      Top             =   525
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7065
      Left            =   2250
      ScaleHeight     =   7035
      ScaleWidth      =   9900
      TabIndex        =   0
      Top             =   0
      Width           =   9930
      Begin VB.Timer Timer5 
         Enabled         =   0   'False
         Interval        =   1000
         Left            =   9360
         Top             =   1080
      End
      Begin VB.Timer Timer4 
         Enabled         =   0   'False
         Interval        =   15000
         Left            =   9360
         Top             =   1680
      End
   End
   Begin VB.Label L_P 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1320
      TabIndex        =   23
      Top             =   7440
      Width           =   375
   End
   Begin VB.Label L_Sup 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   22
      Top             =   7440
      Width           =   375
   End
   Begin VB.Label L_SD 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1320
      TabIndex        =   21
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label L_HP 
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   360
      TabIndex        =   20
      Top             =   5640
      Width           =   375
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "FPS:"
      ForeColor       =   &H00FFFFFF&
      Height          =   315
      Left            =   150
      TabIndex        =   3
      Top             =   1125
      Width           =   1590
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "生命:100"
      BeginProperty Font 
         Name            =   "华文细黑"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   465
      Left            =   75
      TabIndex        =   2
      Top             =   3375
      Width           =   1890
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "速度：5m/s"
      BeginProperty Font 
         Name            =   "华文细黑"
         Size            =   15
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   765
      Left            =   150
      TabIndex        =   1
      Top             =   2175
      Width           =   1515
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command2_Click()
al = -12000
Call 重来按钮_Click
End Sub

Private Sub Command3_Click()

Select Case GMusicPlaying
Case True
GMusic.Pause
GMusicPlaying = False
Case False
GMusic.Play
GMusicPlaying = True
End Select

End Sub





Private Sub Form_Load()

al = -12000
游戏结束评分栏.Top = al
Picture1.Left = 2500
Picture1.Top = 0
Picture1.Width = (Form1.Width - 1750) / 1.2
Picture1.Height = (Form1.Height - 500) / 1.2


加载
Form1.Hide
Form1.Show
Form1.L_HP.Caption = HPNum
Form1.L_SD.Caption = SpeedDownNum
Form1.L_Sup.Caption = SpeedUpNum
Form1.L_P.Caption = ProtectNum
Picture1.Left = 2500
Picture1.Top = 0
Picture1.Width = Form1.Width - 1750
Picture1.Height = Form1.Height + 1000

HP = 100

 Command1_Click

End Sub

Private Sub Command1_Click() '“继续”按钮
Frame1.Visible = False

Do '开始进入游戏循环
    DoEvents '用DoEvents函数，让Windows空出来作别的
    '——————————死了之后————————————————
    If Died = True Then
    If HasFlash = False Then
    effect.Flash 0.9, 0, 0, 2000
    End If
    cam.SetPosition CamX - 100, -70, CamZ
    Form1.Label1.Caption = "速度：" & "0 m/s"
    Form1.Label2.Caption = "生命：0"
    Form1.Label3.Caption = "FPS: 0"
    Timer3.Enabled = False
    Timer5.Enabled = False
   If STcam / 20 > HighScore Then HighScore = Int(STcam / 20)
    总路程.Caption = "总路程: " & STcam / 20 & " m"
    末速度.Caption = "最后的速度: " & Vcam & " m/s"
    奖励金币.Caption = "奖励金币(有速度加成):" & Int((STcam / 20)) & " + " & Int(Vcam) & " x 500 = " & Int(STcam / 20) + Int(Vcam) * 500
    L_HighScore.Caption = "路程 最高记录：" & HighScore & " m"
    游戏结束评分栏.Visible = True
'————————评分栏下降————————————————
    Timer2.Enabled = True
   HasFlash = True
   '——————————记录剩余钱和道具——————————————
    If al >= 1500 Then
    MoneyOwn = MoneyOwn + Int(STcam / 20) + Int(Vcam) * 500
    记录剩余钱与道具
    GoTo n: '不渲染了，让评分屏下来吧
    End If
    
    GoTo d:
   End If
    
f开始
f按键检测
f边界检测
f障碍重置
f回归起点
f碰撞测试
d:
f渲染

SleepTimeLeft = 12 - TV.TimeElapsed 'LOCK住80FPS
If SleepTimeLeft < 0 Then SleepTimeLeft = 0  '运行不够的时间就休眠够，12.5毫秒

Call Sleep(SleepTimeLeft)

Loop Until InputE.IsKeyPressed(TV_KEY_ESCAPE) = True '

n:

If 游戏结束评分栏.Visible = False Then
Frame1.Visible = True
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
Command1.SetFocus
End If


End Sub

Private Sub Command1_GotFocus()
Command1.BackColor = &H80FF80
Command2.BackColor = &HFFFFFF
Command3.BackColor = &HFFFFFF
Command4.BackColor = &HFFFFFF
End Sub


Private Sub Command2_GotFocus()
Command2.BackColor = &H80FF80
Command1.BackColor = &HFFFFFF
Command3.BackColor = &HFFFFFF
Command4.BackColor = &HFFFFFF
End Sub

Private Sub Command3_GotFocus()
Command3.BackColor = &H80FF80
Command1.BackColor = &HFFFFFF
Command2.BackColor = &HFFFFFF
Command4.BackColor = &HFFFFFF
End Sub

Private Sub Command4_GotFocus()
Command4.BackColor = &H80FF80
Command1.BackColor = &HFFFFFF
Command2.BackColor = &HFFFFFF
Command3.BackColor = &HFFFFFF
End Sub


Private Sub Command4_Click()
Form1.Hide
Form2.Label1.Left = 825
Form2.Label2.Left = 3225
Form2.Label3.Left = 5625
Form2.Label4.Left = 8775
Form2.Label5.Left = 7275
Form2.Label3.Caption = "商 店"
Form2.Label3.FontSize = 48
Form2.Label3.Top = 5100
Form2.Picture = LoadPicture(GPath & "bg.jpg")
Form2.Text1 = "请输入购买个数 "
Form2.Text2 = "请输入购买个数 "
Form2.Text3 = "请输入购买个数 "
Form2.Text4 = "请输入购买个数 "
Form2.HP持有数 = "持有数：" & HPNum
Form2.SDown持有数 = "持有数：" & SpeedDownNum
Form2.Sup持有数 = "持有数：" & SpeedUpNum
Form2.Prt持有数 = "持有数：" & ProtectNum
Form2.持有XO币.Caption = "持有XO币：" & MoneyOwn
Form2.Label2.Enabled = False
Form2.Label3.Enabled = False
Form2.Label4.Enabled = False
Form2.Label5.Enabled = False
Form2.Text5 = Form2.Text6
Form2.Show
Set Form2.Picture = Nothing
Form2.Image2.Left = Form2.Width / 2 - Form2.Image2.Width / 2
Form2.Image2.Top = Form2.Height / 2 - Form2.Image2.Height / 2
Form2.Frame1.Left = Form2.Width / 2 - Form2.Frame1.Width / 2
Form2.Frame1.Top = Form2.Height / 2 - Form2.Frame1.Height / 2
Form2.Frame2.Left = Form2.Width / 2 - Form2.Frame2.Width / 2
Form2.Frame2.Top = Form2.Height / 2 - Form2.Frame2.Height / 2
Form2.T_微博.Visible = True
GMusic.Stop_
bgm.Play
Frame1.Visible = False
游戏结束评分栏.Visible = False
al = -12000
游戏结束评分栏.Top = al
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False

Died = False
HasFlash = False

Set TV = Nothing
Set scene = Nothing
Set tex = Nothing
Set InputE = Nothing
Set cam = Nothing
Set body = Nothing
Set curtain = Nothing
Set Atmo = Nothing
For i = 1 To 10
Set sd(i) = Nothing
Set planet(i) = Nothing
Set razer(i) = Nothing
Next
For b = 1 To 100
Set box(b) = Nothing
Next
Set GMusic = Nothing
Set lightE = Nothing
Set Dead_Razer = Nothing
Set effect = Nothing
Set mesh = Nothing
Set Atmo = Nothing
Unload Form1

Form2.Show
Form2.Label2.Enabled = True
Form2.Label3.Enabled = True
Form2.Label4.Enabled = True
Form2.Label5.Enabled = True
SetCursorPos 500, 500
SetCursorPos 501, 501
End Sub


Private Sub Form_Unload(Cancel As Integer)
Call Command4_Click
Cancel = 0
End Sub

Private Sub Timer1_Timer() '歌曲完结判断
If GMusic.Play = False Then
  mNum = mNum + 1
    If mNum = 7 Then mNum = 1
    Set GMusic = Nothing
    Set GMusic = New TVSoundMP3
   GMusic.Load GPath & "bgm/" & mNum & ".mp3"
   GMusic.Play
End If
End Sub

Private Sub Timer2_Timer() '评分栏下降
    al = al + 150
    重来按钮.Enabled = False
   返回按钮.Enabled = False
    游戏结束评分栏.Top = al
    
    If al >= 1500 Then
    重来按钮.Enabled = True
    返回按钮.Enabled = True
    Timer2.Enabled = False
    End If

End Sub

Private Sub Timer3_Timer() '瞄准激光定时激活
GMusic.Volume = -1300
RazerSound.Play
Timer5.Enabled = True
End Sub



Private Sub Timer4_Timer() '无敌防护罩
isBeingProtecting = False
Timer4.Enabled = False
sf_alpha = 0.2
PrtSurface.SetPosition 3000, 0, 0
End Sub

Private Sub Timer5_Timer() '瞄准激光的延迟发射
razer_shoot = True
Timer5.Enabled = False
End Sub

Private Sub 返回按钮_Click()
Form1.Hide
Form2.Label1.Left = 825
Form2.Label2.Left = 3225
Form2.Label3.Left = 5625
Form2.Label4.Left = 8775
Form2.Label5.Left = 7275
Form2.Label3.Caption = "商 店"
Form2.Label3.FontSize = 48
Form2.Label3.Top = 5100
Form2.Picture = LoadPicture(GPath & "bg.jpg")
Form2.Text1 = "请输入购买个数 "
Form2.Text2 = "请输入购买个数 "
Form2.Text3 = "请输入购买个数 "
Form2.Text4 = "请输入购买个数 "
Form2.HP持有数 = "持有数：" & HPNum
Form2.SDown持有数 = "持有数：" & SpeedDownNum
Form2.Sup持有数 = "持有数：" & SpeedUpNum
Form2.Prt持有数 = "持有数：" & ProtectNum
Form2.持有XO币.Caption = "持有XO币：" & MoneyOwn
Form2.Label2.Enabled = False
Form2.Label3.Enabled = False
Form2.Label4.Enabled = False
Form2.Label5.Enabled = False
Set Form2.Picture = Nothing
Form2.Image2.Left = Form2.Width / 2 - Form2.Image2.Width / 2
Form2.Image2.Top = Form2.Height / 2 - Form2.Image2.Height / 2
Form2.Frame1.Left = Form2.Width / 2 - Form2.Frame1.Width / 2
Form2.Frame1.Top = Form2.Height / 2 - Form2.Frame1.Height / 2
Form2.Frame2.Left = Form2.Width / 2 - Form2.Frame2.Width / 2
Form2.Frame2.Top = Form2.Height / 2 - Form2.Frame2.Height / 2
Form2.T_微博.Visible = True
GMusic.Stop_
bgm.Play
Frame1.Visible = False
游戏结束评分栏.Visible = False
al = -12000
游戏结束评分栏.Top = al
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False


Died = False
HasFlash = False

Set TV = Nothing
Set scene = Nothing
Set tex = Nothing
Set InputE = Nothing
Set Dead_Razer = Nothing
Set cam = Nothing
Set body = Nothing
Set curtain = Nothing
Set Atmo = Nothing
For i = 1 To 10
Set sd(i) = Nothing
Set planet(i) = Nothing
Set razer(i) = Nothing
Next
For b = 1 To 100
Set box(b) = Nothing
Next
Set GMusic = Nothing
Set lightE = Nothing
Set effect = Nothing
Set mesh = Nothing
Set Atmo = Nothing
Unload Form1

Form2.Show
Form2.Label2.Enabled = True
Form2.Label3.Enabled = True
Form2.Label4.Enabled = True
Form2.Label5.Enabled = True
SetCursorPos 500, 500
SetCursorPos 501, 501
End Sub

Private Sub 重来按钮_Click()
Vcam = 8
Scam = 0
STcam = 0
RVel = 0
UPVel = 0
HP = 100
al = -12000
Died = False
HasFlash = False
GPath = App.Path & "\data\"
CamX = 300
CamY = 0
CamZ = 0
Died = False
  razerX = CamX + 40000
  razer_shoot = False
  HasWrittenDownZY = False
  Dead_Razer.SetPosition razerX, 0, 0
HasFlash = False
Timer2.Enabled = False
Timer3.Enabled = False
Timer4.Enabled = False
Timer5.Enabled = False
GMusic.Volume = 0
游戏结束评分栏.Visible = False
游戏结束评分栏.Top = -12000
al = -12000
For rz = 1 To 10
razer(rz).SetPosition 1000, 0, 0
Next
effect.FadeIn 5000
Call Command1_Click
End Sub
