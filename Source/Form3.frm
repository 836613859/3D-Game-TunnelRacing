VERSION 5.00
Begin VB.Form Form3 
   BackColor       =   &H008080FF&
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   8235
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12990
   LinkTopic       =   "Form3"
   ScaleHeight     =   8235
   ScaleWidth      =   12990
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   WindowState     =   2  'Maximized
   Begin VB.Label l��� 
      BackStyle       =   0  'Transparent
      Caption         =   "����XO�ң����ٶȼӳɣ���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5100
      TabIndex        =   3
      Top             =   5100
      Width           =   6690
   End
   Begin VB.Label lĩ�ٶ� 
      BackStyle       =   0  'Transparent
      Caption         =   "�����ٶȣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5175
      TabIndex        =   2
      Top             =   4350
      Width           =   6390
   End
   Begin VB.Label l��·�� 
      BackStyle       =   0  'Transparent
      Caption         =   "��·�̣�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   465
      Left            =   5175
      TabIndex        =   1
      Top             =   3525
      Width           =   5715
   End
   Begin VB.Label l���� 
      BackStyle       =   0  'Transparent
      Caption         =   "�� Ϸ �� ��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   42
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1065
      Left            =   5025
      TabIndex        =   0
      Top             =   2025
      Width           =   5340
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
  Dim rtn As Long
  rtn = GetWindowLong(hwnd, GWL_EXSTYLE)
  rtn = rtn Or WS_EX_LAYERED
  SetWindowLong hwnd, GWL_EXSTYLE, rtn
  SetLayeredWindowAttributes hwnd, &H8080FF, 150, LWA_ALPHA 'LWA_COLORKEY���ڿ�
End Sub


