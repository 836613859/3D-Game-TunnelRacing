VERSION 5.00
Begin VB.Form Form2 
   BackColor       =   &H80000012&
   BorderStyle     =   0  'None
   Caption         =   "������"
   ClientHeight    =   11100
   ClientLeft      =   2715
   ClientTop       =   0
   ClientWidth     =   15135
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":324A
   ScaleHeight     =   11100
   ScaleWidth      =   15135
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   WindowState     =   2  'Maximized
   Begin VB.Timer Timer2 
      Interval        =   1000
      Left            =   3840
      Top             =   5520
   End
   Begin VB.TextBox T_΢�� 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "���ߣ�X.X.O.X.X ����΢����http://weibo.com/u/1820770491"
      Top             =   2640
      Width           =   7815
   End
   Begin VB.TextBox Text7 
      Height          =   615
      Left            =   2040
      MultiLine       =   -1  'True
      TabIndex        =   30
      Text            =   "Form2.frx":128B6
      Top             =   3720
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.TextBox Text6 
      Height          =   615
      Left            =   720
      MultiLine       =   -1  'True
      TabIndex        =   29
      Text            =   "Form2.frx":129A7
      Top             =   3840
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   600
      Top             =   5160
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00000000&
      Caption         =   "�̵�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   8415
      Left            =   6960
      TabIndex        =   5
      Top             =   1680
      Visible         =   0   'False
      Width           =   7890
      Begin VB.TextBox Text4 
         Height          =   240
         Left            =   2760
         MaxLength       =   7
         TabIndex        =   20
         Text            =   "�����빺�����"
         Top             =   5625
         Width           =   1515
      End
      Begin VB.TextBox Text3 
         Height          =   240
         Left            =   375
         MaxLength       =   7
         TabIndex        =   19
         Text            =   "�����빺�����"
         Top             =   5640
         Width           =   1515
      End
      Begin VB.TextBox Text2 
         Height          =   240
         Left            =   2700
         MaxLength       =   7
         TabIndex        =   18
         Text            =   "�����빺�����"
         Top             =   2625
         Width           =   1515
      End
      Begin VB.TextBox Text1 
         Height          =   240
         Left            =   360
         MaxLength       =   7
         TabIndex        =   15
         Text            =   "�����빺�����"
         Top             =   2625
         Width           =   1515
      End
      Begin VB.CommandButton Command5 
         BackColor       =   &H0080FFFF&
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   5550
         MaskColor       =   &H00808080&
         Style           =   1  'Graphical
         TabIndex        =   14
         Top             =   7500
         Width           =   2040
      End
      Begin VB.CommandButton Command4 
         BackColor       =   &H0080FFFF&
         Caption         =   "��5000 ���� �޵�״̬15��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   2775
         Style           =   1  'Graphical
         TabIndex        =   13
         Top             =   6000
         Width           =   1680
      End
      Begin VB.CommandButton Command3 
         BackColor       =   &H00C0C0FF&
         Caption         =   "��700 ���� �ٶ�����0.5m/s"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   765
         Left            =   375
         Style           =   1  'Graphical
         TabIndex        =   12
         Top             =   6000
         Width           =   1635
      End
      Begin VB.CommandButton Command2 
         BackColor       =   &H00C0FFC0&
         Caption         =   "��500 ���� �ٶȽ���0.3m/s"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   2700
         Style           =   1  'Graphical
         TabIndex        =   11
         Top             =   3075
         Width           =   1560
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00C0C0FF&
         Caption         =   "��2000 ���� ��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   690
         Left            =   375
         Style           =   1  'Graphical
         TabIndex        =   10
         Top             =   3075
         Width           =   1635
      End
      Begin VB.PictureBox Picture4 
         Height          =   990
         Left            =   600
         Picture         =   "Form2.frx":12E47
         ScaleHeight     =   930
         ScaleWidth      =   930
         TabIndex        =   9
         Top             =   1200
         Width           =   990
      End
      Begin VB.PictureBox Picture3 
         Height          =   990
         Left            =   675
         Picture         =   "Form2.frx":13B6A
         ScaleHeight     =   930
         ScaleWidth      =   930
         TabIndex        =   8
         Top             =   4125
         Width           =   990
      End
      Begin VB.PictureBox Picture2 
         Height          =   990
         Left            =   2925
         Picture         =   "Form2.frx":149A8
         ScaleHeight     =   930
         ScaleWidth      =   930
         TabIndex        =   7
         Top             =   1200
         Width           =   990
      End
      Begin VB.PictureBox Picture1 
         Height          =   990
         Left            =   2925
         Picture         =   "Form2.frx":15A34
         ScaleHeight     =   930
         ScaleWidth      =   930
         TabIndex        =   6
         Top             =   4200
         Width           =   990
      End
      Begin VB.Label Prt������ 
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   2925
         TabIndex        =   23
         Top             =   5250
         Width           =   1140
      End
      Begin VB.Label Sup������ 
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   675
         TabIndex        =   22
         Top             =   5250
         Width           =   1140
      End
      Begin VB.Label SDown������ 
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   2850
         TabIndex        =   21
         Top             =   2325
         Width           =   1140
      End
      Begin VB.Label HP������ 
         BackStyle       =   0  'Transparent
         Caption         =   "��������"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   240
         Left            =   600
         TabIndex        =   17
         Top             =   2325
         Width           =   1140
      End
      Begin VB.Image Image1 
         Height          =   450
         Left            =   4680
         Picture         =   "Form2.frx":172B4
         Top             =   360
         Width           =   840
      End
      Begin VB.Label ����XO�� 
         BackStyle       =   0  'Transparent
         Caption         =   "����XO�ң�0"
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000F&
         Height          =   315
         Left            =   1560
         TabIndex        =   16
         Top             =   450
         Width           =   3090
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00000000&
      Height          =   6375
      Left            =   7680
      TabIndex        =   24
      Top             =   2760
      Visible         =   0   'False
      Width           =   6735
      Begin VB.TextBox Text5 
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   28
         Text            =   "Form2.frx":1C69D
         Top             =   720
         Width           =   6255
      End
      Begin VB.CommandButton Command6 
         Caption         =   "����"
         Height          =   495
         Left            =   2760
         TabIndex        =   27
         Top             =   5760
         Width           =   1335
      End
      Begin VB.OptionButton Option2 
         BackColor       =   &H80000007&
         Caption         =   "����˵��"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   1560
         TabIndex        =   26
         Top             =   360
         Width           =   1335
      End
      Begin VB.OptionButton Option1 
         BackColor       =   &H80000007&
         Caption         =   "��Ϸ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000E&
         Height          =   255
         Left            =   240
         TabIndex        =   25
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "˵ ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   48
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1140
      Left            =   7275
      TabIndex        =   4
      Top             =   6750
      Width           =   3165
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "�˳���Ϸ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   48
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   1215
      Left            =   8775
      TabIndex        =   3
      Top             =   8625
      Width           =   4590
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "�� ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   48
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   990
      Left            =   5625
      TabIndex        =   2
      Top             =   5100
      Width           =   3465
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "��ʼ��Ϸ"
      BeginProperty Font 
         Name            =   "����"
         Size            =   48
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   990
      Left            =   3225
      TabIndex        =   1
      Top             =   3375
      Width           =   4365
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "�� �� �� �"
      BeginProperty Font 
         Name            =   "���Ĳ���"
         Size            =   63.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000F&
      Height          =   1365
      Left            =   825
      TabIndex        =   0
      Top             =   960
      Width           =   7965
   End
   Begin VB.Image Image2 
      Height          =   12000
      Left            =   2160
      Picture         =   "Form2.frx":1C91B
      Top             =   240
      Width           =   19200
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public WithEvents SMovie As TVMovie
Attribute SMovie.VB_VarHelpID = -1


Private Sub Command1_Click() '������
On Error GoTo er:

Q = Int(Text1)
If MoneyOwn - Q * 2000 >= 0 And HPNum + Q <= 20 Then
HPNum = HPNum + Q
HP������.Caption = "��������" & HPNum
MoneyOwn = MoneyOwn - Q * 2000
����XO��.Caption = "����XO�ң�" & MoneyOwn
��¼ʣ��Ǯ�����
Else
  If MoneyOwn - Q * 2000 < 0 Then
  MsgBox "XO�Ҳ���...", vbCritical, "ûǮ..."
  GoTo ex:
  End If
  If HPNum + Q > 20 Then
  MsgBox "����Ʒ����20����", vbCritical, "����"
  GoTo ex:
  End If
End If

GoTo ex:
er: '�����˵Ļ�����
MsgBox "����ȷ����һ����������", vbCritical, "����"
ex: '�˳�

End Sub

Private Sub Command2_Click() '�ٶȼ�
On Error GoTo er:
Q = Int(Text2)
If MoneyOwn - Q * 500 >= 0 And SpeedDownNum + Q <= 99 Then
SpeedDownNum = SpeedDownNum + Q
SDown������.Caption = "��������" & SpeedDownNum
MoneyOwn = MoneyOwn - Q * 500
����XO��.Caption = "����XO�ң�" & MoneyOwn
��¼ʣ��Ǯ�����
Else
  If MoneyOwn - Q * 500 < 0 Then
  MsgBox "XO�Ҳ���...", vbCritical, "ûǮ..."
  GoTo ex:
  End If
  If SpeedDownNum + Q > 99 Then
  MsgBox "����Ʒ����99����", vbCritical, "����"
  GoTo ex:
  End If
End If

GoTo ex:
er:
MsgBox "����ȷ����һ����������", vbCritical, "����"
ex:

End Sub

Private Sub Command3_Click() '�ٶȼ�
On Error GoTo er:
Q = Int(Text3)
If MoneyOwn - Q * 700 >= 0 And SpeedUpNum + Q <= 50 Then
SpeedUpNum = SpeedUpNum + Q
Sup������.Caption = "��������" & SpeedUpNum
MoneyOwn = MoneyOwn - Q * 700
����XO��.Caption = "����XO�ң�" & MoneyOwn
��¼ʣ��Ǯ�����
Else
  If MoneyOwn - Q * 700 < 0 Then
  MsgBox "XO�Ҳ���...", vbCritical, "ûǮ..."
  GoTo ex:
  End If
  If SpeedUpNum + Q > 50 Then
  MsgBox "����Ʒ����50����", vbCritical, "����"
  GoTo ex:
  End If
End If

GoTo ex:
er:
MsgBox "����ȷ����һ����������", vbCritical, "����"
ex:

End Sub

Private Sub Command4_Click() '�޵�
On Error GoTo er:
Q = Int(Text4)
If MoneyOwn - Q * 5000 >= 0 And ProtectNum + Q <= 10 Then
ProtectNum = ProtectNum + Q
Prt������.Caption = "��������" & ProtectNum
MoneyOwn = MoneyOwn - Q * 5000
����XO��.Caption = "����XO�ң�" & MoneyOwn
��¼ʣ��Ǯ�����
Else
  If MoneyOwn - Q * 5000 < 0 Then
  MsgBox "XO�Ҳ���...", vbCritical, "ûǮ..."
  GoTo ex:
  End If
  If ProtectNum + Q > 10 Then
  MsgBox "����Ʒ����10����", vbCritical, "����"
  GoTo ex:
  End If
End If

GoTo ex:
er:
MsgBox "����ȷ����һ����������", vbCritical, "����"
ex:

End Sub






Private Sub Command5_Click()
Frame1.Visible = False
Label2.Enabled = True
Label3.Enabled = True
Label4.Enabled = True
Label5.Enabled = True
Text1 = "�����빺����� "
Text2 = "�����빺����� "
Text3 = "�����빺����� "
Text4 = "�����빺����� "
End Sub



Private Sub Command6_Click()
Frame2.Visible = False
End Sub

Private Sub Command7_Click()
SMovie.PlayMovie
End Sub

Private Sub Form_Load()


If HasPlayedMovie = False Then
Set SMovie = New TVMovie
With SMovie
        .FileName = App.Path & "\data\s"
        .Balance = 0
        .Volume = 0
        .PlayRate = 1
        .VideoWindowLeft = 0
        .VideoWindowTop = 0
        .VideoWindowHeight = 0
        .VideoWindowWidth = 0
        .VideoWindowStyle = eFullScreenStyle
        .Initialize Form2.hwnd
End With
SMovie.PlayMovie 0
Form2.Show
Do
DoEvents
Loop Until SMovie.Duration <= SMovie.Position
HasPlayedMovie = True
Unload Form2
Form2.Show
Text5 = Text6
Set Form2.Picture = Nothing
Image2.Left = Form2.Width / 2 - Image2.Width / 2
Image2.Top = Form2.Height / 2 - Image2.Height / 2
Frame1.Left = Form2.Width / 2 - Frame1.Width / 2
Frame1.Top = Form2.Height / 2 - Frame1.Height / 2
Frame2.Left = Form2.Width / 2 - Frame2.Width / 2
Frame2.Top = Form2.Height / 2 - Frame2.Height / 2

SetCursorPos 500, 500
SetCursorPos 600, 600
���زֿ��뱳������
End If

'If HasPlayedMovie = True Then



End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = &H8000&
Label3.ForeColor = &H8000&
Label4.ForeColor = &H8000&
Label5.ForeColor = &H8000&
sound.Stop_
End Sub

Private Sub Frame1_Click()
Text1 = "�����빺����� "
Text2 = "�����빺����� "
Text3 = "�����빺����� "
Text4 = "�����빺����� "
End Sub


Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label2.ForeColor = &HFF& Then GoTo ex:
Label2.ForeColor = &HFF&
Label3.ForeColor = &H8000&
Label4.ForeColor = &H8000&
Label5.ForeColor = &H8000&
sound.Play
ex:
End Sub


Private Sub Label3_Click()
Frame1.Visible = True
Frame2.Visible = False
Label2.Enabled = False
Label3.Enabled = False
Label4.Enabled = False
Label5.Enabled = False
End Sub

Private Sub Label3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label3.ForeColor = &HFF& Then GoTo ex:
Label3.ForeColor = &HFF&
Label2.ForeColor = &H8000&
Label4.ForeColor = &H8000&
Label5.ForeColor = &H8000&
sound.Play
ex:
End Sub


Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label4.ForeColor = &HFF& Then GoTo ex:
Label4.ForeColor = &HFF&
Label2.ForeColor = &H8000&
Label3.ForeColor = &H8000&
Label5.ForeColor = &H8000&
sound.Play

ex:
End Sub

Private Sub Label5_Click()
Frame2.Visible = True
Frame1.Visible = False
End Sub

Private Sub Label5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Label5.ForeColor = &HFF& Then GoTo ex:
Label5.ForeColor = &HFF&
Label2.ForeColor = &H8000&
Label3.ForeColor = &H8000&
Label4.ForeColor = &H8000&
sound.Play
ex:
End Sub


Private Sub Label2_Click()
bgm.Stop_
Image2.Left = 50000
Frame1.Visible = False
Frame2.Visible = False
T_΢��.Visible = False
Label1.Left = 50000
Label2.Left = 50000

Form2.Label3.Caption = "Loading...."
Label3.Left = Form2.Width / 2
Label3.Top = Form2.Height / 2
Label3.FontSize = 20
Label3.ForeColor = &HFFFFFF
Label4.Left = 50000
Label5.Left = 50000
Timer1.Enabled = False

Form1.Show

End Sub
Private Sub label4_click()

f�������
End Sub

Private Sub Option1_Click()
Text5 = Text6
End Sub

Private Sub Option2_Click()
'Text5 = "����Ϸ������" & Chr(13) & Chr(10) & "�������� ����λ��" & Chr(13) & Chr(10) & "W ���˵���֮�����" & Chr(13) & Chr(10) & "S ���˵���֮�����" & Chr(13) & Chr(10) & "A ���˵���֮����Ѫ" & Chr(13) & Chr(10) & "D ���˵���֮���޵�15��" & Chr(13) & Chr(10) & "PageDown ��Ϸ�����󻻸�" & Chr(13) & Chr(10) & "PageUp ��Ϸ����ǰ����" & Chr(13) & Chr(10) & "Esc ��Ϸ�е����˵�"
Text5 = Text7
End Sub
'����������������������������������TEXT��������������������������������
Private Sub Text1_Click()
If Len(Text1) > 4 Then Text1 = ""
End Sub
Private Sub Text2_Click()
If Len(Text2) > 4 Then Text2 = ""
End Sub
Private Sub Text3_Click()
If Len(Text3) > 4 Then Text3 = ""
End Sub
Private Sub Text4_Click()
If Len(Text4) > 4 Then Text4 = ""
End Sub


Private Sub Timer1_Timer()
bgm.Play
End Sub

Private Sub Timer2_Timer()
If SMovie.Position >= SMovie.Duration Then SMovie.StopMovie
Timer2.Enabled = False
End Sub
