VERSION 5.00
Begin VB.Form ReadyFrm 
   ClientHeight    =   7320
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   10476
   LinkTopic       =   "Form1"
   ScaleHeight     =   7320
   ScaleWidth      =   10476
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.Timer Timer1 
      Left            =   576
      Top             =   504
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "���̕�����A"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   36
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   720
      Index           =   1
      Left            =   3360
      TabIndex        =   2
      Top             =   1320
      Width           =   4080
   End
   Begin VB.Label Label2 
      Alignment       =   2  '��������
      AutoSize        =   -1  'True
      Caption         =   "System not ready"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   25.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   516
      Left            =   3720
      TabIndex        =   1
      Top             =   3360
      Width           =   3828
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "�ُ탊�Z�b�g�������Ă�������"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   36
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   720
      Index           =   0
      Left            =   1080
      TabIndex        =   0
      Top             =   2160
      Width           =   8484
   End
End
Attribute VB_Name = "ReadyFrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim lViewFlg      '�O�̉�ʔԍ�
Private Sub Form_Load()
  DispCenter Me
  lViewFlg = ViewFlg      '�O�̉�ʔԍ�
  'ViewFlg = 2             '��ʔԍ�
  FrmMenuFlg = True                   '���j���[���甲����Ƃ�false
  Timer1.Interval = 100
  Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
  Dim ms$
  Timer1.Enabled = False
  Do
    DoEvents
    If SystemReadyChk() = 1 Then Exit Do
    Label2.Caption = MsgChk
  Loop
  '
  Select Case lViewFlg
  Case 1
    PGM_Menu.Show
    Unload Me
  Case 2 '���`�i�V���O���j
    NQD70_SC.Show
    Unload Me
  Case 3  '���`�i�e�X�g�j
    LS21_TC.Show
    Unload Me
  Case 4  'I O �`�F�b�N
    IOChk.Show
  Case 5  '�X�P�[���ύX
    LS21_GphScale.Show
  Case 6  '�ǂݏo��
  Case 7  '������
  Case 8  'edit
    MYEdit.Show
    Unload Me
  Case Else
    PGM_Menu.Show
    Unload Me
  End Select
  
End Sub
