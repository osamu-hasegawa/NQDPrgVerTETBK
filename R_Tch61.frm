VERSION 5.00
Begin VB.Form R_Tch61 
   BackColor       =   &H00E0E0E0&
   Caption         =   "Form1"
   ClientHeight    =   4284
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   6276
   LinkTopic       =   "Form1"
   ScaleHeight     =   4284
   ScaleWidth      =   6276
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   240
      Top             =   960
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�L�����Z��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   90
      Width           =   1236
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�I��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   1
      Left            =   1410
      TabIndex        =   0
      Top             =   90
      Width           =   1236
   End
   Begin VB.Label Label2 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   1
      Left            =   3180
      TabIndex        =   5
      Top             =   1200
      Width           =   1500
   End
   Begin VB.Label Label2 
      Alignment       =   1  '�E����
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   3180
      TabIndex        =   4
      Top             =   840
      Width           =   1500
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�ʒu�i���j�F"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   8
      Left            =   1560
      TabIndex        =   3
      Top             =   1200
      Width           =   1545
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�ʒu�i��j�F"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Index           =   9
      Left            =   1560
      TabIndex        =   2
      Top             =   885
      Width           =   1545
   End
End
Attribute VB_Name = "R_Tch61"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' R_Tch61�@��۸���
   
    'update 2002.8.10   ROZ( ) �˂����Đ��`�p���Ұ��֕ύX



Dim lViewFlg      '�O�̉�ʔԍ�

Private Sub Command2_Click(Index As Integer)
  Select Case Index
  Case 0  '�L�����Z��
    
  Case 1  '�I��
    Unload Me
    PGM_Menu.Show
  
  End Select
End Sub

Private Sub Form_Load()
  DispCenter Me
  Timer1.Enabled = False
  lViewFlg = ViewFlg      '�O�̉�ʔԍ�
  ViewFlg = 9             '��ʔԍ�
  FrmMenuFlg = True       '���j���[���甲����Ƃ�false
  Me.Show
  SetData
  Timer1.Enabled = True
End Sub

Private Sub Timer1_Timer()
Dim up%, dw%, md%, ve%
Dim pos!
  DioInput 10, dw
  DioInput 9, up
  DioInput 11, md
  DioInput 12, ve
  pos = r_z()
  If up = 1 Then
    Label2(0).Caption = Format(pos, "0.000")
'    If md = 1 Then roz(0) = pos
  End If
  If dw = 1 Then
    Label2(1).Caption = Format(pos, "0.000")
'    If md = 1 Then roz(1) = pos
  End If
  
  
End Sub

Private Sub SetData()
  ServoON
  C870OrgVelSet   '/* ���_�p���x�ݒ� */
  Label2(4).Caption = "���_�o�����s"
  genten
  Ready_Wait
  Label2(4).Caption = "���_�o������"
  C870ManVelSet   '/* �蓮�p���x�ݒ� */
  '/* �J�E���^�Ƀ[������������ */
  C870AdrInit       '�`�c�c�q�d�r�r �h�m�h�s�`�k�h�y�d �b�n�l�l�`�m�c
  C870CntPreSet 0   '�b�n�t�m�s�d�q �o�q�d�r�d�s �b�n�l�l�`�m�c
  
End Sub
Private Sub genten()
'--------------
  C870Genten
End Sub

