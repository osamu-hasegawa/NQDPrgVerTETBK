VERSION 5.00
Begin VB.Form NQD70_SC 
   Appearance      =   0  '�ׯ�
   BackColor       =   &H00C0C0C0&
   Caption         =   "�A�����`"
   ClientHeight    =   8532
   ClientLeft      =   132
   ClientTop       =   420
   ClientWidth     =   11844
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   ScaleHeight     =   8532
   ScaleWidth      =   11844
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Scr.Copy"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   260
      Index           =   2
      Left            =   10960
      Style           =   1  '���̨���
      TabIndex        =   124
      Top             =   8160
      Width           =   760
   End
   Begin VB.CommandButton Command1 
      Caption         =   "����"
      Height          =   252
      Left            =   11350
      TabIndex        =   123
      Top             =   2280
      Width           =   490
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  '�ׯ�
      BackColor       =   &H00E0E0E0&
      Caption         =   "�^��"
      ForeColor       =   &H80000008&
      Height          =   1572
      Left            =   10250
      TabIndex        =   110
      Top             =   2760
      Width           =   1575
      Begin VB.Line Line3 
         X1              =   0
         X2              =   1560
         Y1              =   1320
         Y2              =   1320
      End
      Begin VB.Label Label13 
         Alignment       =   2  '��������
         BorderStyle     =   1  '����
         Caption         =   "Label13"
         Enabled         =   0   'False
         Height          =   240
         Index           =   8
         Left            =   1200
         TabIndex        =   119
         Top             =   1320
         Visible         =   0   'False
         Width           =   390
      End
      Begin VB.Label Label13 
         Alignment       =   2  '��������
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '����
         Caption         =   "Label13"
         Enabled         =   0   'False
         Height          =   240
         Index           =   7
         Left            =   120
         TabIndex        =   118
         Top             =   1320
         Visible         =   0   'False
         Width           =   405
      End
      Begin VB.Label Label13 
         Alignment       =   2  '��������
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   1  '����
         Caption         =   "Label13"
         Height          =   240
         Index           =   6
         Left            =   600
         TabIndex        =   117
         Top             =   1320
         Width           =   405
      End
      Begin VB.Label Label13 
         Alignment       =   2  '��������
         BackColor       =   &H00FFFFC0&
         BorderStyle     =   1  '����
         Caption         =   "Label13"
         Height          =   240
         Index           =   5
         Left            =   120
         TabIndex        =   116
         Top             =   960
         Width           =   405
      End
      Begin VB.Label Label13 
         Alignment       =   2  '��������
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '����
         Caption         =   "Label13"
         Height          =   240
         Index           =   4
         Left            =   120
         TabIndex        =   115
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label13 
         Alignment       =   2  '��������
         BackColor       =   &H008080FF&
         BorderStyle     =   1  '����
         Caption         =   "Label13"
         Height          =   240
         Index           =   3
         Left            =   600
         TabIndex        =   114
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label13 
         Alignment       =   2  '��������
         BackColor       =   &H00C0C0FF&
         BorderStyle     =   1  '����
         Caption         =   "Label13"
         Height          =   240
         Index           =   2
         Left            =   1080
         TabIndex        =   113
         Top             =   240
         Width           =   405
      End
      Begin VB.Label Label13 
         Alignment       =   2  '��������
         BackColor       =   &H00C0E0FF&
         BorderStyle     =   1  '����
         Caption         =   "Label13"
         Height          =   240
         Index           =   1
         Left            =   1095
         TabIndex        =   112
         Top             =   600
         Width           =   405
      End
      Begin VB.Label Label13 
         Alignment       =   2  '��������
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  '����
         Caption         =   "Label13"
         Height          =   240
         Index           =   0
         Left            =   600
         TabIndex        =   111
         Top             =   960
         Width           =   405
      End
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "5����~"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   0
      Left            =   2520
      Style           =   1  '���̨���
      TabIndex        =   106
      Top             =   100
      Width           =   500
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00FFFFC0&
      Caption         =   "�ۉ���~"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   9
      Left            =   3240
      Style           =   1  '���̨���
      TabIndex        =   95
      Top             =   100
      Width           =   500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "PlotDataSave"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   5
      Left            =   120
      TabIndex        =   57
      Top             =   600
      Width           =   1428
   End
   Begin VB.ListBox List1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1140
      Left            =   2088
      TabIndex        =   77
      Top             =   1080
      Width           =   8124
   End
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   0
      Top             =   120
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�����\�[�N"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   400
      Index           =   8
      Left            =   1800
      Style           =   1  '���̨���
      TabIndex        =   59
      Top             =   100
      Width           =   500
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�^�󓞒B"
      Enabled         =   0   'False
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
      Index           =   4
      Left            =   0
      TabIndex        =   56
      Top             =   1320
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.CommandButton Command2 
      BackColor       =   &H00C0FFC0&
      Caption         =   "V �G�f�B�^"
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
      Index           =   3
      Left            =   120
      Style           =   1  '���̨���
      TabIndex        =   54
      Top             =   1080
      Width           =   1440
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   1920
      Top             =   4200
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '�ׯ�
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      ForeColor       =   &H80000008&
      Height          =   5136
      Left            =   1800
      ScaleHeight     =   5112
      ScaleWidth      =   8376
      TabIndex        =   8
      Top             =   2240
      Width           =   8400
      Begin VB.ListBox List3 
         Appearance      =   0  '�ׯ�
         BackColor       =   &H00800000&
         ForeColor       =   &H000000FF&
         Height          =   924
         Left            =   5160
         TabIndex        =   122
         Top             =   50
         Width           =   3156
      End
      Begin VB.ListBox List2 
         Appearance      =   0  '�ׯ�
         BackColor       =   &H00800000&
         ForeColor       =   &H00FFFFFF&
         Height          =   744
         Left            =   0
         TabIndex        =   93
         Top             =   240
         Width           =   4932
      End
      Begin VB.Label Label10 
         Appearance      =   0  '�ׯ�
         BackColor       =   &H00800000&
         BackStyle       =   0  '����
         Caption         =   "Label10"
         ForeColor       =   &H00FFFFFF&
         Height          =   225
         Left            =   0
         TabIndex        =   94
         Top             =   0
         Width           =   7455
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '�_��
         Index           =   7
         X1              =   6696
         X2              =   6696
         Y1              =   0
         Y2              =   6436
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '�_��
         Index           =   6
         X1              =   5040
         X2              =   5040
         Y1              =   0
         Y2              =   6436
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '�_��
         Index           =   5
         X1              =   3348
         X2              =   3348
         Y1              =   0
         Y2              =   6436
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '�_��
         Index           =   4
         X1              =   1656
         X2              =   1656
         Y1              =   0
         Y2              =   6436
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '�_��
         Index           =   3
         X1              =   0
         X2              =   8352
         Y1              =   1010
         Y2              =   1010
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '�_��
         Index           =   2
         X1              =   0
         X2              =   8352
         Y1              =   2030
         Y2              =   2030
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '�_��
         Index           =   1
         X1              =   0
         X2              =   8352
         Y1              =   3060
         Y2              =   3060
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '�_��
         Index           =   0
         X1              =   0
         X2              =   8352
         Y1              =   4090
         Y2              =   4090
      End
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
      Left            =   120
      Style           =   1  '���̨���
      TabIndex        =   5
      Top             =   120
      Width           =   1440
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "-"
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
      Index           =   4
      Left            =   10200
      TabIndex        =   121
      Top             =   75
      Width           =   135
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
      Left            =   9740
      TabIndex        =   120
      Top             =   75
      Width           =   420
   End
   Begin VB.Label Label12 
      Alignment       =   2  '��������
      BackColor       =   &H00FFC0FF&
      Caption         =   "Label12"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   2
      Left            =   10250
      TabIndex        =   109
      Top             =   2520
      Width           =   1572
   End
   Begin VB.Label Label12 
      BackColor       =   &H00FFC0FF&
      Caption         =   "Label12"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Index           =   1
      Left            =   10250
      TabIndex        =   108
      Top             =   2280
      Width           =   1572
   End
   Begin VB.Label Label12 
      Alignment       =   2  '��������
      BackColor       =   &H00FFC0FF&
      Caption         =   "Label12"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Index           =   0
      Left            =   10250
      TabIndex        =   107
      Top             =   2040
      Width           =   1572
   End
   Begin VB.Label Label11 
      Alignment       =   2  '��������
      BackColor       =   &H00C0C0C0&
      Caption         =   "label11"
      Height          =   216
      Index           =   9
      Left            =   10240
      TabIndex        =   105
      Top             =   7160
      Width           =   200
   End
   Begin VB.Label Label11 
      Alignment       =   2  '��������
      BackColor       =   &H00C0C0C0&
      Caption         =   "label11"
      Height          =   204
      Index           =   8
      Left            =   10240
      TabIndex        =   104
      Top             =   6884
      Width           =   200
   End
   Begin VB.Label Label11 
      Alignment       =   2  '��������
      BackColor       =   &H00C0C0C0&
      Caption         =   "label11"
      Height          =   216
      Index           =   7
      Left            =   10240
      TabIndex        =   103
      Top             =   6596
      Width           =   200
   End
   Begin VB.Label Label11 
      Alignment       =   2  '��������
      BackColor       =   &H00C0C0C0&
      Caption         =   "label11"
      Height          =   216
      Index           =   6
      Left            =   10240
      TabIndex        =   102
      Top             =   6320
      Width           =   200
   End
   Begin VB.Label Label11 
      Alignment       =   2  '��������
      BackColor       =   &H00C0C0C0&
      Caption         =   "label11"
      Height          =   216
      Index           =   5
      Left            =   10240
      TabIndex        =   101
      Top             =   6044
      Width           =   200
   End
   Begin VB.Label Label11 
      Alignment       =   2  '��������
      BackColor       =   &H00C0C0C0&
      Caption         =   "label11"
      Height          =   204
      Index           =   4
      Left            =   10240
      TabIndex        =   100
      Top             =   5756
      Width           =   200
   End
   Begin VB.Label Label11 
      Alignment       =   2  '��������
      BackColor       =   &H00C0C0C0&
      Caption         =   "label11"
      Height          =   216
      Index           =   3
      Left            =   10240
      TabIndex        =   99
      Top             =   5480
      Width           =   200
   End
   Begin VB.Label Label11 
      Alignment       =   2  '��������
      BackColor       =   &H00C0C0C0&
      Caption         =   "label11"
      Height          =   216
      Index           =   2
      Left            =   10240
      TabIndex        =   98
      Top             =   5204
      Width           =   200
   End
   Begin VB.Label Label11 
      Alignment       =   2  '��������
      BackColor       =   &H00C0C0C0&
      Caption         =   "label11"
      Height          =   204
      Index           =   1
      Left            =   10240
      TabIndex        =   97
      Top             =   4916
      Width           =   200
   End
   Begin VB.Label Label11 
      Alignment       =   2  '��������
      BackColor       =   &H00C0C0C0&
      Caption         =   "label11"
      Height          =   216
      Index           =   0
      Left            =   10240
      TabIndex        =   96
      Top             =   4640
      Width           =   200
   End
   Begin VB.Label Label8 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Label8"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   4000
      TabIndex        =   92
      Top             =   90
      Width           =   1120
   End
   Begin VB.Label Label9 
      Caption         =   "  Z3�␳"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Index           =   1
      Left            =   11170
      TabIndex        =   91
      Top             =   4400
      Width           =   580
   End
   Begin VB.Label Label9 
      Caption         =   "  �s�W��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   200
      Index           =   0
      Left            =   10460
      TabIndex        =   90
      Top             =   4400
      Width           =   612
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   0
      Left            =   11200
      TabIndex        =   89
      Top             =   4640
      Width           =   540
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   1
      Left            =   11200
      TabIndex        =   88
      Top             =   4916
      Width           =   540
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   2
      Left            =   11200
      TabIndex        =   87
      Top             =   5204
      Width           =   540
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   3
      Left            =   11200
      TabIndex        =   86
      Top             =   5480
      Width           =   540
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   4
      Left            =   11200
      TabIndex        =   85
      Top             =   5756
      Width           =   540
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   5
      Left            =   11200
      TabIndex        =   84
      Top             =   6044
      Width           =   540
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   6
      Left            =   11200
      TabIndex        =   83
      Top             =   6320
      Width           =   540
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   7
      Left            =   11200
      TabIndex        =   82
      Top             =   6596
      Width           =   540
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   8
      Left            =   11200
      TabIndex        =   81
      Top             =   6884
      Width           =   540
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   9
      Left            =   11200
      TabIndex        =   80
      Top             =   7160
      Width           =   540
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   216
      Index           =   1
      Left            =   11200
      TabIndex        =   79
      Top             =   7500
      Width           =   540
   End
   Begin VB.Label Label7 
      Caption         =   "Label7"
      Height          =   216
      Index           =   0
      Left            =   10500
      TabIndex        =   78
      Top             =   7500
      Width           =   540
   End
   Begin VB.Label Label5 
      Caption         =   "cc3-2"
      Height          =   252
      Index           =   6
      Left            =   10320
      TabIndex        =   76
      Top             =   1440
      Width           =   1380
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   9
      Left            =   10500
      TabIndex        =   75
      Top             =   7160
      Width           =   540
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   8
      Left            =   10500
      TabIndex        =   74
      Top             =   6884
      Width           =   540
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   7
      Left            =   10500
      TabIndex        =   73
      Top             =   6596
      Width           =   540
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   6
      Left            =   10500
      TabIndex        =   72
      Top             =   6320
      Width           =   540
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   5
      Left            =   10500
      TabIndex        =   71
      Top             =   6044
      Width           =   540
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   4
      Left            =   10500
      TabIndex        =   70
      Top             =   5756
      Width           =   540
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   3
      Left            =   10500
      TabIndex        =   69
      Top             =   5480
      Width           =   540
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   216
      Index           =   2
      Left            =   10500
      TabIndex        =   68
      Top             =   5204
      Width           =   540
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      Height          =   216
      Index           =   1
      Left            =   10500
      TabIndex        =   67
      Top             =   4916
      Width           =   540
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   8.4
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   216
      Index           =   0
      Left            =   10500
      TabIndex        =   66
      Top             =   4640
      Width           =   540
   End
   Begin VB.Label Label5 
      Caption         =   "cc3"
      Height          =   252
      Index           =   5
      Left            =   8640
      TabIndex        =   65
      Top             =   4560
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label Label5 
      Caption         =   "cc2"
      Height          =   252
      Index           =   4
      Left            =   8640
      TabIndex        =   64
      Top             =   4080
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label Label5 
      Caption         =   "cc1"
      Height          =   252
      Index           =   3
      Left            =   8640
      TabIndex        =   63
      Top             =   3480
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label Label5 
      Caption         =   "ct2"
      Height          =   252
      Index           =   2
      Left            =   8640
      TabIndex        =   62
      Top             =   3120
      Visible         =   0   'False
      Width           =   1452
   End
   Begin VB.Label Label5 
      Caption         =   "ct1"
      Height          =   252
      Index           =   1
      Left            =   10320
      TabIndex        =   61
      Top             =   1100
      Width           =   1380
   End
   Begin VB.Label Label5 
      Caption         =   "cp1"
      Height          =   252
      Index           =   0
      Left            =   10320
      TabIndex        =   60
      Top             =   1780
      Width           =   1380
   End
   Begin VB.Label Label2 
      BackColor       =   &H80000005&
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
      Height          =   252
      Index           =   14
      Left            =   6720
      TabIndex        =   58
      Top             =   7800
      Width           =   4980
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
      Index           =   13
      Left            =   11040
      TabIndex        =   55
      Top             =   72
      Width           =   660
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�R�}���h�F"
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
      Index           =   11
      Left            =   120
      TabIndex        =   53
      Top             =   8160
      Width           =   1290
   End
   Begin VB.Label Label2 
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
      Index           =   12
      Left            =   1428
      TabIndex        =   52
      Top             =   8160
      Width           =   5040
   End
   Begin VB.Label Label2 
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
      Index           =   11
      Left            =   6720
      TabIndex        =   51
      Top             =   8160
      Width           =   4140
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�V���b�g���F"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   9
      Left            =   8400
      TabIndex        =   50
      Top             =   120
      Width           =   1275
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�T�C�N���^�C���F"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   9.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   195
      Index           =   8
      Left            =   8400
      TabIndex        =   49
      Top             =   480
      Width           =   1695
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
      Index           =   9
      Left            =   10340
      TabIndex        =   48
      Top             =   75
      Width           =   420
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
      Index           =   8
      Left            =   10200
      TabIndex        =   47
      Top             =   480
      Width           =   1500
   End
   Begin VB.Label Label2 
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
      Index           =   7
      Left            =   1440
      TabIndex        =   46
      Top             =   7800
      Width           =   5040
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�R�}���h�F"
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
      Index           =   2
      Left            =   120
      TabIndex        =   45
      Top             =   7800
      Width           =   1296
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   6
      Left            =   8400
      TabIndex        =   44
      Top             =   780
      Width           =   3312
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   5
      Left            =   5280
      TabIndex        =   43
      Top             =   3360
      Visible         =   0   'False
      Width           =   4872
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9.6
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   4
      Left            =   3480
      TabIndex        =   42
      Top             =   3480
      Visible         =   0   'False
      Width           =   3432
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "���`��ԁF"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   228
      Index           =   1
      Left            =   2040
      TabIndex        =   41
      Top             =   3240
      Visible         =   0   'False
      Width           =   1212
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "(��)"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   204
      Index           =   31
      Left            =   9360
      TabIndex        =   40
      Top             =   7560
      Width           =   468
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�o�ߎ���"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   204
      Index           =   30
      Left            =   7236
      TabIndex        =   39
      Top             =   7560
      Width           =   912
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   27
      X1              =   10200
      X2              =   10200
      Y1              =   7380
      Y2              =   7488
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   26
      X1              =   8520
      X2              =   8520
      Y1              =   7380
      Y2              =   7488
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   25
      X1              =   6840
      X2              =   6840
      Y1              =   7380
      Y2              =   7488
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   24
      X1              =   5160
      X2              =   5160
      Y1              =   7380
      Y2              =   7488
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   23
      X1              =   3480
      X2              =   3480
      Y1              =   7380
      Y2              =   7488
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   29
      Left            =   9930
      TabIndex        =   38
      Top             =   7485
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   28
      Left            =   8355
      TabIndex        =   37
      Top             =   7485
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   27
      Left            =   6660
      TabIndex        =   36
      Top             =   7485
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   26
      Left            =   4965
      TabIndex        =   35
      Top             =   7485
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   25
      Left            =   3270
      TabIndex        =   34
      Top             =   7485
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   210
      Index           =   24
      Left            =   1650
      TabIndex        =   33
      Top             =   7485
      Width           =   375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   22
      X1              =   1800
      X2              =   1800
      Y1              =   7380
      Y2              =   7488
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00000000&
      Index           =   21
      X1              =   10200
      X2              =   1800
      Y1              =   7380
      Y2              =   7380
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�^���x"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   204
      Index           =   23
      Left            =   1212
      TabIndex        =   32
      Top             =   1620
      Width           =   684
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "(��)"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   204
      Index           =   22
      Left            =   1200
      TabIndex        =   31
      Top             =   1872
      Width           =   468
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      Index           =   20
      X1              =   1620
      X2              =   1764
      Y1              =   2220
      Y2              =   2220
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      Index           =   19
      X1              =   1620
      X2              =   1764
      Y1              =   3248
      Y2              =   3248
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      Index           =   18
      X1              =   1620
      X2              =   1824
      Y1              =   4276
      Y2              =   4276
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      Index           =   17
      X1              =   1620
      X2              =   1764
      Y1              =   5304
      Y2              =   5304
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      Index           =   16
      X1              =   1620
      X2              =   1764
      Y1              =   6332
      Y2              =   6332
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000C000&
      Index           =   15
      X1              =   1620
      X2              =   1764
      Y1              =   7350
      Y2              =   7350
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      Index           =   14
      X1              =   1776
      X2              =   1776
      Y1              =   2220
      Y2              =   7384
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "####"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   204
      Index           =   21
      Left            =   1212
      TabIndex        =   30
      Top             =   2124
      Width           =   480
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   210
      Index           =   20
      Left            =   1290
      TabIndex        =   29
      Top             =   3148
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   210
      Index           =   19
      Left            =   1290
      TabIndex        =   28
      Top             =   4170
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   216
      Index           =   18
      Left            =   1320
      TabIndex        =   27
      Top             =   5209
      Width           =   372
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   210
      Index           =   17
      Left            =   1290
      TabIndex        =   26
      Top             =   6232
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000C000&
      Height          =   210
      Index           =   16
      Left            =   1290
      TabIndex        =   25
      Top             =   7260
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�^����"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   204
      Index           =   15
      Left            =   516
      TabIndex        =   24
      Top             =   1620
      Width           =   684
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "(kg)"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   204
      Index           =   14
      Left            =   612
      TabIndex        =   23
      Top             =   1872
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   13
      X1              =   1005
      X2              =   1149
      Y1              =   2220
      Y2              =   2220
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   12
      X1              =   1005
      X2              =   1149
      Y1              =   3248
      Y2              =   3248
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   11
      X1              =   1005
      X2              =   1149
      Y1              =   4276
      Y2              =   4276
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   10
      X1              =   1005
      X2              =   1149
      Y1              =   5304
      Y2              =   5304
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   9
      X1              =   1005
      X2              =   1149
      Y1              =   6332
      Y2              =   6332
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   8
      X1              =   1005
      X2              =   1149
      Y1              =   7350
      Y2              =   7350
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   7
      X1              =   1152
      X2              =   1152
      Y1              =   2220
      Y2              =   7360
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   13
      Left            =   645
      TabIndex        =   22
      Top             =   2120
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   12
      Left            =   645
      TabIndex        =   21
      Top             =   3148
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   11
      Left            =   645
      TabIndex        =   20
      Top             =   4170
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   10
      Left            =   645
      TabIndex        =   19
      Top             =   5209
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   9
      Left            =   645
      TabIndex        =   18
      Top             =   6232
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   210
      Index           =   8
      Left            =   645
      TabIndex        =   17
      Top             =   7260
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "���W"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   204
      Index           =   7
      Left            =   36
      TabIndex        =   16
      Top             =   1620
      Width           =   456
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "(mm)"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   204
      Index           =   6
      Left            =   48
      TabIndex        =   15
      Top             =   1872
      Width           =   480
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   6
      X1              =   390
      X2              =   534
      Y1              =   2220
      Y2              =   2220
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   5
      X1              =   390
      X2              =   534
      Y1              =   3248
      Y2              =   3248
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   4
      X1              =   390
      X2              =   534
      Y1              =   4276
      Y2              =   4276
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   3
      X1              =   390
      X2              =   534
      Y1              =   5304
      Y2              =   5304
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   2
      X1              =   390
      X2              =   534
      Y1              =   6332
      Y2              =   6332
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   1
      X1              =   390
      X2              =   534
      Y1              =   7350
      Y2              =   7350
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   0
      X1              =   540
      X2              =   540
      Y1              =   2220
      Y2              =   7360
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   5
      Left            =   30
      TabIndex        =   14
      Top             =   2120
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   4
      Left            =   30
      TabIndex        =   13
      Top             =   3148
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   3
      Left            =   30
      TabIndex        =   12
      Top             =   4170
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   2
      Left            =   30
      TabIndex        =   11
      Top             =   5209
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   1
      Left            =   30
      TabIndex        =   10
      Top             =   6232
      Width           =   375
   End
   Begin VB.Label Label3 
      Alignment       =   1  '�E����
      BackStyle       =   0  '����
      Caption         =   "###"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FFFF&
      Height          =   210
      Index           =   0
      Left            =   30
      TabIndex        =   9
      Top             =   7260
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�R�����g�F"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   0
      Left            =   1920
      TabIndex        =   7
      Top             =   780
      Width           =   1215
   End
   Begin VB.Label Label2 
      BackColor       =   &H8000000E&
      Caption         =   "Label2"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   10.8
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   3
      Left            =   3240
      TabIndex        =   6
      Top             =   780
      Width           =   4930
   End
   Begin VB.Label Label2 
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
      Index           =   2
      Left            =   3960
      TabIndex        =   4
      Top             =   480
      Width           =   4215
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "����t�@�C�����F"
      BeginProperty Font 
         Name            =   "�l�r �S�V�b�N"
         Size            =   11.4
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   225
      Index           =   7
      Left            =   1950
      TabIndex        =   3
      Top             =   480
      Width           =   1935
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "��"
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
      Index           =   5
      Left            =   7968
      TabIndex        =   2
      Top             =   72
      Width           =   276
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
      Left            =   6840
      TabIndex        =   1
      Top             =   72
      Width           =   1008
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "���莞�ԁF"
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
      Index           =   3
      Left            =   5520
      TabIndex        =   0
      Top             =   84
      Width           =   1272
   End
End
Attribute VB_Name = "NQD70_SC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    NQD70_SC
'            update: 2002.6.28 s.f  private sub cal_pid�@�폜
'            update: 2002.6.28 s.f difftime�@��������
'            update: 2002.7.10 s.f "DC","HC" �V�K�ǉ�
'            update: 2002.8.10 s.f roz(0),roz(1)��˓����`�����Ұ��� max.180
'            update: 2002.8.15 s.f Veditcol �ǉ�
'            update: 2002.8.18 s.f �^�N�g�^�C���\�� int(stime/60)��
'                                  "HC" �C�����A������
'                                  "DC" ���@���`�񐔁@�߂��ii_s=i_s-1)
'
'            update: 2002.8.22 s.f ���W���@���F��
'            update: 2002.8.24 s.f �b��ύX�@�uVEdit���@����K������v
'            update: 2002.8.25 s.f ���`�񐔁@save�@�@InitDTsave�@���@���`��ֈړ�
'            update: 2002.8.29 s.f cp,ct,cc�f�[�^�\��'
'            update: 2002.9.06 s.f ���`�񐔁@�\���@idcflg�ǉ�
'            update: 2002.9.26 s.f ic(i)=10 �Ł@�I�����f�@�Ɂ@����
'            update: 2002.10.1 s.f �����䃂�[�h�Q�ցA�@CtlDisp  'DioOut 12,1  �ʒu���� '  02.10.1 �ǉ�
'            update: 2002.10.1 s.f ������@�G���[�\���@Label2(4)����Label2(3)�֕ύX
'            update: 2002.10.2 s.f ������X�^�[�g���ԕ\��
'            update: 2002.10.5 s.f �^�C���A�b�v���[�`���������i�����Ĕ�ё΍�j
'            update: 2002.10.5 s.f ���ԕ\���ύX
'            update: 2002.10.9 KYOCERA �^�C�}�[�����A�^�C���A�b�v�A�R�����g�\���A���ԕ\���ύX
'            update: 2002.10.12 s.f ��ѱ��߂̐�����@goto���@�ύX
'            update: 2002.10.16 KYOCERA ��ѱ��ߏ��� <9 �� istend �ɕύX
'            update: 2002.10.16 KYOCERA ��ѱ��߂Ŏ��̽ï�ߒǉ�
'            update: 2002.10.17 KYOCERA ���_���A��ɏ��񌴓_���A�����׸�gOrgStartFlg��ON
'            update: 2002.10.17 KYOCERA ��ѱ��ߏ��� <istend �� 10 �ɕύX
'            update: 2002.10.26 s.f ������@�G���[�\���@Label2(3)����Label2(5)�֕ύX
'            �@�@�@�@�@�@�@�@�@�@s.f cc3-cc2�\���ǉ�
'                                   SR�@�̏����ύX�@0.1�b�ɂP������ݸ�
'            update: 2002.11.28 s.f �I����t�E�����@�ύX�@�i�����\�ɂ���j
'            update: 2002.12.03 s.f ���`�L�^�̕\���E�f�B�X�N�L�^�@�ǉ�
'            update: 2002.12.05 s.f ���`�L�^�̕\���E�f�B�X�N�L�^�@�C��
'            update: 2003.03.22 s.f CT�R�}���h�@��L�����@ct=  -> ct_temp(  ��
'            update: 2003.07.10 HND �A���[���\�����́@���`�v���O�������s
'                                  frmerr_sign, FbiDio, LS21_SC
'            update: 2004. 3. 8 s.f. LS21_SC �ύX�@���`�����䃂�[�h�@�f�V�f�ǉ��@�i�㎲�Փ˔���t�j
'                                    RecEmgDTsave ����~���b�Z�[�W�̕ۑ�
'
'            update: 2004. 3.12 s.f.  ���x�w�ߓd���@�\��
'            update: 2004. 4.23 s.f.  timeup�Ł@����~
'            update: 2004. 5. 5 s.f   ���x�W���A�����␳���[�`���@�ǉ�  PGM_KTD,My_lib,MYEDIT, LS21_SC, LS21_TC
'            update: 2004.5.12  s.f   PGM_KTD�@"���ް�۰"�΍�@�@wTm0!,wTm1!  global��,  LS21_SC�Ɓ@LS21_TC ����@dim�폜
'            update: 2004.5.17  s.f   'S'����ށ@�o�O�΍�
'            update: 2004.5.18  s.f    T�W���\��
'            update: 2004.8.17  s.f   ���ް�۰"�΍�  p(ist0)��pp��  �h�F�h�����̍s�𖳂���
'                                     List1.Enabled = True or False �ǉ�
'            update: 2004.8.27 - 10.30  s.f   T�W���֐��ύX�A�@�@�u�c�b�@�O�v�R�}���h�@���`�O�Ɍ^�ݔۃ`�F�b�N�Z���T�[�̃`�F�b�N�@�\�ǉ�
'            update: 2005. 5.25 s.f    Version No�\���ǉ�
'            update: 2005. 7.18 s.f    �������ԁ@���ϒl�\��
'            update: 2005. 7.25 s.f   �������Ԑ���f�o�b�O    List2.Enabled = True or False �ǉ�
'            update: 2005. 9.27 s.f    �ۉ���~���[�h�ǉ�  ���`�I�����@���������炸�ɕۉ����Ē�~
'            update: 2005. 9.28 s.f   T�W���@�\���F�ύX
'            update: 2005.11. 4 s.f �@ LS21_SC�@�\���ύX�B���x����d���\���폜�BT�W���AZ�R�␳�\�����ύX�A�@�������Ԑ���o�O�C��
'            update: 2005.11.22 s.f   Melec C-870 counter����o�O�C���@�R���y�A�J�E���^�l�Z�b�g���@�������]�@�@setcm1
'                                     C870sts(3) ����@�o�O�C���A�E���f�[�^�����ύX
'            update: 2005.11.23 s.f   11/22 �ύX�̃o�O�C���@���`������@�uC870sts�@reset����܂Ł@�ǂݔ�΂��v���@����
'�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@��ʉ��\���@�V���v�����@�i�X�s�[�h�ቺ�h�~�ׁ̈j
'            update: 2005.11.26 s.f   ���ׂẮ@function�@�Ɂ@�^�錾������@�@�@overflow�΍�
'            update: 2005.12.17 s.f   Do-Loop �O�́@DoEvent�폜 OverFlow �΍� s.f.
'                                     �R�}���h�́@evtime�@��荞�݂��@�R�}���h�J�n���֕ύX
'�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@DC�R�}���h�@LA�R�}���h�@�ă`�F�b�N�C��
'�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�A���O�R�}���h�@evtime�@�Ɓ@fintime�@�\�L����ւ�
'            update: 2005.12.23 s.f
'            update: 2006. 2.18 s.f
'            update: 2006. 3. 3 s.f  edit �g�p���@do�@loop���甲����
'�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@DC����ނց@fintime=timer�@���@�ݒu
'            update: 2006. 4.14 s.f  on error goto,  sts as long
'            update: 2006. 4.15 s.f  error �\��
'            update: 2006. 5. 9 s.f  O.F.error �\���@������@end3�@�ǉ�,  tstime=0#
'            update: 2006. 5.14 s.f �@r_pres()�́@DoEvents �@ for�̊O�ֈړ��@s.f  ���̂���������
'�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@  ���ׂĔ����Ɓ@LS_TC�@�v���O�����\������iLS_SC�́@OK)�f
'            update: 2006. 5.15 s.f  5���ԕۉ���~�@�ǉ�
'            update: 2006. 5.18 s.f �@r_pres()�́@DoEvents �@�폜�A�@�hJ"�A�P�b��1��@Doevents�@�ǉ�
'                                    ����~�@�\���ǉ�
'            update: 2006. 7.12 s.f  �������Ԏ��������@�f�L���f��
'
'       Ver.3.33R_061221 2006.12.21 s.f  LS-33���@�Ή��@�@VacuumON�AVacuumOFF�@��p�~�ASeikeiON,SeikeiOFF�V�݁@DO3�@���蓖�ĕύX
'       Ver.3.33R_070827 2007.08.27 s.f  ����~���́@���u�ǉ�
'       Ver.3.33R_070927 2007.09.27 s.f  Z�␳�@�w�肵��������No.�ց@�ł���悤�ɂ���
'       Ver.3.33R_071112 2007.11.13 s.f  �u�����\�[�N�v�����A�@�u1�񐬌`�venable=False��
'       Ver.3.33R_071119 2007.11.19 s.f  �������Ԑ���@�o�O�C���iedit���A�f�[�^�p���j�A���ϒlAND�ŐV�l�Ł@�X�V�����
'       Ver.3.33R_071120 2007.11.20 s.f  �o�O�C���A�@�󐬌`-�r�o�@�ǉ��A�@�A�����`�ĊJ�@�ǉ�
'       Ver.3.33R_071121 2007.11.21 s.f  ��������@���ϒl�v�Z�@����̉������ԁ@�d��2.0��
'       Ver.3.33R_071122 2007.11.22 s.f  �^���@�\���o�O�C��
'       Ver.3.33R_071127 2007.11.27 s.f  �^���@�\���|�C���^�[���֕ύX
'       Ver.3.33R_071210 2007.12.10 s.f  �I�����@T�W�����i�[���ā@�I������l�ύX�i�@save�@�ǉ��@�j
' --- NQD
'       Ver.NQD080312 2008.2.12 s.f  NewQD���`�@�@Ver.
'       Ver.NQD_71_081205 2008.12. 5 s.f  ���`���̕\���@�������D�����@���A�������ԁA�b���@�A���[��
'       Ver.NQD_71_090817 s.f  SystemNotReady�@�Q��`�F�b�N�A�A���[���\���@�P�b�ɂP��`�F�b�N�X�V�ցA
'       Ver.NQD_71_100306 2010.3. 6 s.f.  ����|�C���^�[����@�o�O�C��
'�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�V���b�g���@���`�L�������m�F���́@if i_s >0 then ... endif ���폜
'           '
'       Ver.NQD_71_100405 2010. 4. 5 s.f. timeup�����@�@skip������@LongTime��to(ist0)�֕ύX
'�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@����|�C���^�[�Y���̏C��100306�̃o�O���
'       Ver.NQD_71_100407 2010.4. 7 s.f. timeup���� skip���� �o�O�C���F�@���肩��u������R�}���h�@�X�̎��͏����v
'�@�@�@ Ver.NQD_71_101124�@2010.11.24 s.f. �@���x�ݒ�@T_keisu_cset�i�j ���@ntemp(jsub),otemp(ksub)����폜�B�@���ˉ��x�v�ł͂Ȃ��A�M�d�΂̂���T�W���𔽉f�����Ȃ��B
'�@�@�@ Ver.NQD_71_120624�@2012.06.24 s.f.�@������1,3,7�̏ꍇ�@z���B���X�^�[�g���Ƀ`�F�b�N�ǉ�
'�@�@�@ Ver.NQD_71_130423  2013. 4.23.s.f.  �����щ����i30���ȏ�\�ցjResDt�̌��@2000��12000�i��12000�b�j��
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
Option Explicit
Dim lGphNo%
Dim lGphNo0%
Dim EditFlg As Long
Dim lViewFlg      '�O�̉�ʔԍ�
Dim NextView%
Dim NextViewBUp%  'NextView�̓��ebackup
Dim lDtSaveFlg%   '�f�[�^�ۑ�
Dim idcflg%(0 To 3)        ' DC�t���O�@�`��=1�@�^�L=0
Dim SokuCor!(0 To 1)  '�����\�[�N�^�C���̃R�}���h�t�̐F
Dim TKatBackCol!(0 To 1)  '�������ԕ␳�@��������@�\����backColor
Dim lEmgFlg As Long       '����~
Dim iflghoonStop As Long, iHoonStopNo As Long  '�ۉ���~�t���O�A�ۉ���~�񐔃J�E���^�[
Dim iflg5Stop As Long    '5���ԕۉ���~�t���O
Dim iHoteikanryou As Long  '�ۉ���~�@�m�F�t���O
Dim iflgSCopy As Integer   ' ScreenCopy �t���O
'
'�X�N���[���̃X�i�b�v�V���b�g���N���b�v�{�[�h�ɕۑ��y�ш���@�@�ϐ��錾���@�@�i273�j '

Private Declare Sub keybd_event Lib "user32.dll" _
        (ByVal bVk As Byte, ByVal bScan As Byte, _
         ByVal dwFlags As Long, ByVal dwExtraInfo As Long)

Private Const VK_SNAPSHOT = &H2C            'PrintScreen �L�[(P1051)
Private Const VK_LMENU = &HA4               'Alt�L�[
Private Const KEYEVENTF_KEYUP = &H2         '�L�[�̓A�b�v���
Private Const KEYEVENTF_EXTENDEDKEY = &H1   '�X�L�����R�[�h�͊g���R�[�h
'
Private Sub Command1_Click()
    If iflghoonStop = True Then
     iHoteikanryou = 1
    End If
End Sub

Private Sub Command2_Click(Index As Integer)
Select Case Index
'Case 0  '�L�����Z��
'  lGphNo = 0
'  MoniGraph Me.Picture1, 0, lGphNo
Case 1  '�I��
   If FrmMenuFlg = True Then
          FrmMenuFlg = False          '�I����t
          NextViewBUp = NextView
          NextView = 1
          Command2(1).BackColor = CmndColon(1)
    Else
          FrmMenuFlg = True           '�I����t����
          NextView = NextViewBUp
          Command2(1).BackColor = CmndColoff(1)
  End If
Case 2
'''�A�N�e�B�u�E�C���h�E���N���b�v�{�[�h�ɃR�s�[�������B�@True �ɐݒ�
    Select Case iflgSCopy
        Case 0
            iflgSCopy = 1          'ScreenCopy�@1 ��t�@�L�����Ă܂��͐擪���Ă�copy
            Command2(2).BackColor = CmndColon(1)    ' on 1=red
        Case 1
            iflgSCopy = 2      'ScreenCopy�@2 ��t�@�L�����Ď��̂�copy
           Command2(2).BackColor = CmndColon(9)    ' on 9=pink
        Case 2
            iflgSCopy = 0       'ScreenCopy�@1 ��t����
            Command2(2).BackColor = CmndColoff(0)
    End Select
'
''''�A�N�e�B�u�E�C���h�E���N���b�v�{�[�h�ɃR�s�[�������B�@True �ɐݒ�
'  If iflgSCopy = True Then
'          iflgSCopy = False          'ScreenCopy�@��t����
'          Command2(2).BackColor = CmndColoff(0)
'    Else
'          iflgSCopy = True      'ScreenCopy�@��t
'         Command2(2).BackColor = CmndColon(1)    ' on 1=red
'  End If
'
'Case 2  '�O���t�ĕ`��
'  lGphNo = lGphNo + 100
'  MoniGraph Me.Picture1, 0, lGphNo
''
Case 3                        'edit�@�́@'02/8�b��ύX(s.f)
  If EditFlg = True Then
          EditFlg = False          '�G�f�B�^�N������
          Command2(3).BackColor = CmndColoff(3)
    Else
          EditFlg = True      '�G�f�B�^�N��
          Command2(3).BackColor = CmndColon(1)   ' 1=red
  End If
'
Case 4      '�^�󓞒B
  gVumFlg = 1                       '�^�󓞒B=1
Case 5      '"Save" ;�f�[�^�Z�[�u
'  lDtSaveFlg = True
  
    If lDtSaveFlg = True Then
          lDtSaveFlg = False          '�f�[�^�Z�[�u�@��t����
          Command2(5).BackColor = CmndColoff(1)    ' off gray
          Command2(5).Caption = "Save"
    Else
          lDtSaveFlg = True           '�f�[�^�Z�[�u�@��t
          Command2(5).BackColor = CmndColon(1)   ' on 1= red
          Command2(5).Caption = "DataSave��"
  End If
'
Case 8      '�����\�[�N�^�C��
  If lSokuFlg = True Then
          lSokuFlg = False          '�����\�[�N�^�C���@��t����
          Command2(8).BackColor = SokuCor(0)
    Else
          lSokuFlg = True           '�����\�[�N�^�C���@��t
          Command2(8).BackColor = SokuCor(1)
  End If
Case 9     '�ۉ���~  ���`�I�����@���������炸�ɕۉ����Ē�~
  If iflghoonStop = True Then
          iHoteikanryou = 1
          iflghoonStop = False          '�ۉ���~�@��t����
          Command2(9).BackColor = CmndColoff(9)
    Else
          iflghoonStop = True      '�ۉ���~�@��t
          iHoteikanryou = 0
          Command2(9).BackColor = CmndColon(1)    ' on 1=red
          iflg5Stop = False        '5���ԕۉ���~�@��t����
          Command2(0).BackColor = CmndColoff(0)
  End If
  If (KataChk() < 4) Then  '�^������
          iflghoonStop = False          '�ۉ���~�@��t����
          Command2(9).BackColor = CmndColoff(9)
  End If
Case 0     '5���ԕۉ���~
  If iflg5Stop = True Then
          iflg5Stop = False          '5���ԕۉ���~�@��t����
          Command2(0).BackColor = CmndColoff(0)
    Else
          iflg5Stop = True      '5���ԕۉ���~�@��t
          Command2(0).BackColor = CmndColon(1)    ' on 1=red
          iflghoonStop = False  '�ۉ���~�@��t����
          Command2(9).BackColor = CmndColoff(9)
  End If
  If (KataChk() < 4) Then  '�^������
          iflg5Stop = False          '5���ԕۉ���~�@��t����
          Command2(0).BackColor = CmndColoff(0)
  End If
'
End Select
DoEvents
End Sub

Private Sub SetData()
  Label2(0) = Format(ptime, "###0")  '���莞��
  Label2(2) = gcoxFlName             '����t�@�C����
  Label2(3) = hcomm(2)               '�R�����g
' -----------------------------------
  DispGphScale
End Sub

Private Sub Form_Load()
  DispCenter Me
  NQD70_SC.Caption = NQD70_SC.Caption + "     " + versionNo
  Me.Top = 0
  SokuCor(0) = &H8000000F     '�����\�[�N�^�C���̃R�}���h�t�̐F
  SokuCor(1) = &HFF&          '�����\�[�N�^�C���̃R�}���h�t�̐F �����ꂽ�Ƃ�
  lDtSaveFlg = False      '�f�[�^�ۑ�
'  lSokuFlg = False        '�����\�[�N�^�C��   GPM_KTD��   2007.11.19
  If lSokuFlg = False Then
          Command2(8).BackColor = SokuCor(0)
    Else
          Command2(8).BackColor = SokuCor(1)
  End If
  lViewFlg = ViewFlg      '�O�̉�ʔԍ�
  ViewFlg = 2             '��ʔԍ�
  FrmMenuFlg = True       '���j���[���甲����Ƃ�false
  EditFlg = False        '�G�f�B�^�N������
  Command2(1).BackColor = CmndColoff(1)     '�I���R�}���h�t�̐F
  Command2(3).BackColor = CmndColoff(3)     'V�G�f�B�g�̃R�}���h�t�̐F
  Command2(9).BackColor = CmndColoff(9)     '�ۉ���~�R�}���h�t�̐F
    TKatBackCol!(0) = &H8000000F      '��������@�n�e�e�̂Ƃ�
    TKatBackCol!(1) = &HC0C0FF      '��������@�n�m�̂Ƃ�
    lEmgFlg = False         '����~
  SetData
  Timer1.Enabled = True
  iflghoonStop = False
  iHoonStopNo = 0
End Sub


Private Sub DispGphScale()
Dim i%, p%, max!, min!, def!, dev%
  '
  GphXSet           '���Ԏ��̎��Ԃ��Z�b�g
  '
  dev = 5
  '
  min = InitDat(1)  '�O���t�X�P�[�����W (Min)
  max = InitDat(2)  '�O���t�X�P�[�����W (Max)
  def = (max - min) / dev
  p = 0
  For i = 0 To 5
    Label3(p + i).Caption = Format(min + def * i, "0")
  Next i
  min = InitDat(3)  '�O���t�X�P�[���^���� (Min)
  max = InitDat(4)  '�O���t�X�P�[���^���� (Max)
  def = (max - min) / dev
  p = 8
  For i = 0 To 5
    Label3(p + i).Caption = Format(min + def * i, "0")
  Next i
  min = InitDat(5)  '�O���t�X�P�[���^���x (Min)
  max = InitDat(6)  '�O���t�X�P�[���^���x (Max)
  def = (max - min) / dev
  p = 16
  For i = 0 To 5
    Label3(p + i).Caption = Format(min + def * i, "0")
  Next i
  min = InitDat(7)  '�O���t�X�P�[���o�ߎ��� (Min)
  max = InitDat(8)  '�O���t�X�P�[���o�ߎ��� (Max)
  def = (max - min) / dev
  p = 24
  For i = 0 To 5
    Label3(p + i).Caption = Format(min + def * i, "0")
  Next i
'
End Sub
Private Sub Timer1_Timer()
  Timer1.Enabled = False
  LS21S_MAIN
End Sub
Public Sub LS21S_MAIN()
Dim i%, j%, js%, l%, ist0%, ist1%, iflg%, isflg%, itu%
Dim ied%, ips%, i_s%, I_s0%, irei%, r_ch%, ix%, ix0%, iy%, isp%, i_s_do%
Dim stime%, ii%, iii%, istend%
Dim ie02%, ie03%, ie04%, ituflg%, iSRcount%, iki%, ikii%
Dim ie%, ie0%, ie1%, ie2%, ie3%, ie4%, ie5%, iflghsmsg%
'Dim ie%, ie0%, ie1%, ie2%, ie3%, ie4%, ie5%, iFlg_hijyou%, iflghsmsg%�@�@�@' 090817 iFLG_Hijyou�@���@Global��
Dim m_l%, sv%, zch%
Dim ivd%, id_0%, id_1%, id_2%
Dim ct_dummy!, iz3%, itc%, ict%, ikat%
Dim idmy%, ch%, hdt%, flindex%, imax%, sts1%, sts2%, ch1%, ch2%
Dim sts As Long                                     '2006.4.14
Dim it_ts%, i_ts%
Dim dmy$, sdt$, c$, com$, tdate$, ttime$, kjdisp$
'Dim sdt1$, sdt2$, sdt3$�@�@�@2006.4.14�@global he
Dim isub As Long, jsub As Long, ksub As Long, lsub As Long
Dim flg As Long, cnt As Long
Dim iwt!, S_StartTime!
Dim sdata!    '  05.11.26 s.s. overflow �΍�
Dim ndata!, mdata!, odata!, ntemp!, mtemp!, otemp!, ntemp0!, mtemp0!, otemp0!, htemp!
Dim imachi!, hs5_fintime!, hs5_sttime!, hs5_difft!, hs5_diffTold!
Dim st!, ev!, sev!, fin!, it!, it0!         '/* ���ԗp�f�[�^ */
Dim btemp!(0 To 4), bposi!, bpre! '/* ���x�@�ʒu�@���� �̑O�f�[�^ */
Dim stTime!, evtime!, sevTime!, mTime!, tsTime!, endTime!   ' 2009.8.17 fintime global ��
'Dim stTime!, evtime!, fintime!, sevTime!, mTime!, tsTime!, endTime!
Dim dt!(0 To 7), adFlg As Long
Dim diTime!, diTime1!, diTime2!, diTimeSR!, pdt!, pp!, pml!
Dim x1dt!, x2dt!, pos!
Dim r_z_now!, r_z_ave!, r_z_dum!(0 To 180)    ' /* 2002.7.10�@�ǉ��@�˓����`�@*/
Dim epsilon!
Dim cp_z!, cc_time!(0 To 3), ct_temp!(0 To 2)   ' CP , CT �p
Dim ct_t!(0 To 10)
Dim avekatJ!(0 To 10), katJ!
'Dim kaatsuJ!(0 To 10, 0 To 5), avekatJ!(0 To 10), kjdisp$, katJ!, ikat%
Dim zclear!
Dim idum%, iidum%       ' 090803 tsuika
Dim tudiffTime!
Dim iSento_flg%         ' �擪�_�~�[�׸�
Dim zzz!    ' 2013.4.6 �������B���́@�y���W�l  SP7  180901
'
 On Error GoTo errHandler:
' ---  init  val-----------------
  ppos = "SC"   'NQD70_SC  ���݈ʒu
  ips = 1
    If Saikaiflg = True Then
            i_s = 0                     '�ĊJ���́A���񂩂�J�E���g
        Else
            i_s = -1
    End If
'  i_s = -1            '���`��
'  iz3 = 3            '�@Z3�@�́@index�l�@Z(ist0)    07.9.27  ��ŃZ�b�g�@�@iz3=Z3_HoseiCont(2)
  iFlg_hijyou = 0      ' ����~FLG�̏������@0=�ُ�Ȃ�
  For ii = 0 To 3: idcflg(ii) = 0: Next ii
  For ii = 0 To 10: ct_t(ii) = 0: Next ii
  c = "0"
  ivd = 0:   id_0 = 0: id_2 = &H8
  For ii = 1 To 180: r_z_dum(ii) = 0#: Next ii
  For i = 0 To 5: For ii = 0 To 10: kaatsuJ(ii, i) = 0#: Next ii: Next i
  For ii = 0 To 10: avekatJ(i) = 0#: Next ii
  Label10.Caption = "  No   SL   Ave.   0   -1   -2   -3   New-T Old-T"
  tsTime = 0#
'
  Label12(0).Visible = False
  Label12(1).Visible = False
  Label12(2).Visible = False
  Command1.Visible = False
  iflgSCopy = 0
'
'----------------------- �A�����`���C���v���O����
  C870Stop
  ServoON       '/* �T�[�{���� */
  CtlDisp       '�ʒu����
  TrnsReqOFF    '�����˗��M��OFF
  SeikeiON         '���`ON�@�A�����͂P�񐬌`��
'/***********     �گ��@C-853�{�[�h�����ݒ�@�@�@*************/
'/* SPEC INITIALIZE CMD OUT */
'/* �J�E���^�{�[�h�̏����ݒ� */
  InitDat(10) = 0
'/* ������ڰľ�ĺ���� */
  C870AccRate
'/* ���x�ݒ� */
  C870LSPDSet 300    '/* 300 pps 0.066mm/sec */
'/* �f�B���[�^�C���ݒ� */
  C870DelayTime
  rstcm1   '  compareter reset
'/***********     �گ��@C-853�{�[�h�����ݒ�@�I��  *************/
'/* �`�s�b���x���Z�b�g */
'/* ���{�b�g�f�[�^�̃t���b�s�[����̓ǂ݂Ƃ� */
  rozFileLoad
'
'/* ���`�f�[�^�ۑ��t�@�C���̍쐬�@*/
  RecDtSave0 InitDat(11)
'
'
  it_ts = Int(roz(1))   ' 10       '/* �˂����ĒB���@�����@���ς���� */
  epsilon = roz(0)    ' 0.0005   '/* �˓��@���e���@�@mm�@�@*/
    i_s_do = -1   ' Do Loop �́@��   '�@���`�@Do�@Loop(�{�̂�Do Loop�j�̉񐔁@�@�@�@edit �ŃL�����Z������Ȃ��悤�Ɂ@�����ֈړ��B 2007.11.26
    kataNoPnt = 0  ' �^No �́@�|�C���^�[�@�����ݒ�
'
'-------------------------------------------------------------------------------------
st:             '  Loop�@�P�@�@�i�ŊO���[�v�j
  If ied = 2 Then GoTo st2:             '  ���̕��@�C�ɂȂ�I�I�@ied=2�@�́@�����I�I�@�@edit�̎��́Aied=1�@�@����ȊO�́Aied=0
'  ---�@2007.11.27�@�ǉ��@kataNo�\��  �X�V
    For iii = 0 To katamax
        kataNoHyj(iii) = kataNo(iii)
        kataNoHyj(iii + katamax + 1) = kataNo(iii)
        kataNoHyj(iii + (katamax + 1) * 2) = kataNo(iii)
        kataNoHyj(iii + (katamax + 1) * 3) = kataNo(iii)
    Next iii
'
'/*  ����t�@�C���̃I�[�v�� */
  coxDtRead gcoxFldir & gcoxFlName
  Label2(0).Caption = Format(ptime, "0")
  '/* �O���t�B�b�N��ʂ̏����� */
  InitDat(8) = ptime  '�O���t�X�P�[���o�ߎ��� (Max)
  SetData
  lGphNo0 = 0
  lGphNo = 0
  MoniGraph Me.Picture1, lGphNo0, lGphNo
  For itc = 0 To 9
    Label4(itc).Caption = Format(T_keisu(itc), "0.000")
    Label6(itc).Caption = Format(Z3_Hosei(itc), "0.000")
    If itc < T_keisuCont(0) Then
         Label4(itc).BackColor = T_keisuCol!(1)
         Label6(itc).BackColor = T_keisuCol!(1)
         Label11(itc).Caption = kataNo(itc)
'         Label11(itc).Caption = itc + 1
       Else
         Label4(itc).BackColor = T_keisuCol!(0)
         Label6(itc).BackColor = T_keisuCol!(0)
         Label11(itc).Caption = " "
    End If
    If (iflgKataTorF(itc) = False) Then
         Label4(itc).BackColor = T_keisuCol!(4)
         Label6(itc).BackColor = T_keisuCol!(4)
    End If
  Next itc
  If (katCflag = True) Then
         Label7(0).BorderStyle = 1  '  �g�L��
         Label7(1).BorderStyle = 1  '  �g�L��
    Else
         Label7(0).BorderStyle = 0  '  �g�Ȃ�
         Label7(1).BorderStyle = 0  '  �g�Ȃ�
  End If
''/* �\�����M���x�ݒ� */
'/* ���쓮����R�}���h�̃t�@�C������̓ǂݎ�� */
  i = 0
  Do
    sdt = Right("     " & Format(i, "0"), 4)
    sdt = sdt & "  " & Right("     " & Format(seg_num(i), "0"), 4)
    sdt = sdt & "  " & Right("     " & Format(ic(i), "0"), 4)
    sdt = sdt & "  " & Right("         " & Format(z(i), "0.000"), 7)
    sdt = sdt & "  " & Right("         " & Format(vel(i), "0.0"), 7)
    sdt = sdt & "  " & Right("       " & Format(pres(i), "0"), 6)
    sdt = sdt & "  " & Right("     " & Format(t0(i), "0.0"), 4)
    sdt = sdt & "  " & Right("     " & Format(p(i), "0.0"), 4)
    Label2(12).Caption = sdt
    If pres(i) >= 1000 Then ips = 2    '/* ��ڽ����1ton�ȏ�Ŏ��ύX */
    i = i + 1                          '/*�������`�掞�̃X�P�[���ύX�p*/
    If ic(i - 1) = 9 Then Exit Do
  Loop

  istend = i     '  /* �R�}���h���̍ő�l */
  ic(i) = 10     '  /*  ic(�@)=10 �I���̈Ӗ� */
  'ic(i) = 4     '  /* ����������@���\�t�g�́A�O�`�R������*/
  ic(i + 1) = 10 '  /* ����������@�I���̈Ӗ��@���߉���*/
'
''
'/* �\��̕\�� */
  Label2(2).Caption = gcoxFlName
'/* ���_�o�� */
  Label2(6).Caption = "���_�o�����s"
  genten
  Ready_Wait
  Counter0
  Label2(6).Caption = "���_�o������"
'/* �J�E���^�Ƀ[������������ */
  C870CntPreSet 0   '�b�n�t�m�s�d�q �o�q�d�r�d�s �b�n�l�l�`�m�c
  'InitDat(10) = 0
  pos = r_z()
  GCnt0 = 0
  GCnt1 = 0
'
'
'/* �����^�]�F�� */
  ch1 = 1            '�V�X�e�����f�B�[
  ch2 = 3            '����
  Do
    DoEvents
    If FrmMenuFlg = False Then GoTo eend:            '���j���[���甲����Ƃ�false
'    LS21S_Monitor     '-------------- �s���j�v�ǂ� �^��Ȃ�    '2006.12.21 �폜 s.f
    '
    DioInput ch1, sts1
    DioInput ch2, sts2
    If sts1 = 1 And sts2 = 1 Then Exit Do
  Loop
'/* ���`�v���Z�X�J�n�@�A���O�R�}���h */
  flindex = 0      '����R�}���h�t�@�C���̈ʒu
  Do
    DoEvents
    '-------------- �s���j�v�ǂ�
'    LS21S_Monitor    '2006.12.21 �폜 s.f
    'flindex = flindex + 1
    com = Left(scom(flindex), 1)
    isub = sisub(flindex)
    sdt = Right("    " & scom(flindex), 2)
    sdt = sdt & Right(Space(15) & Format(isub, "0"), 15)
    If (com = "S") Or (com = "L") Then
      jsub = sjsub(flindex)
      ksub = sksub(flindex)
      lsub = slsub(flindex)
      sdt = sdt & Right(Space(15) & Format(jsub, "0"), 15)
      sdt = sdt & Right(Space(15) & Format(ksub, "0"), 15)
      sdt = sdt & Right(Space(15) & Format(lsub, "0"), 15)
    End If
    Label2(7).Caption = sdt
    flindex = flindex + 1
    i = 10
    '
    If ied <> 0 Then GoTo jp0:
    '
    Select Case com
      Case "B"
      Case "N"    '/* ���f�K�X�̒��� */
        If Mid(scom(flindex), 2, 1) = "S" Then
          If isub = 1 Then
'            Label2(4).Caption = "���f�K�X���� DO1"
            N2Open
          End If
          If isub = 0 Then
'            Label2(4).Caption = "���f�K�X��~ DO1"
            N2Close
          End If
        End If
      Case "J"    '/* ���ԑ҂� */
        evtime = Timer

        Do
          fintime = Timer2func
'          fintime = Timer
          DoEvents
          If diffTime(fintime, evtime) >= isub Then Exit Do
        Loop
      Case "K"    '/* ���M */
        Select Case Int(isub)
        Case 1
          HeatON
        Case 0
          HeatOFF
        End Select
      Case "S"    '/* �`�s�b���x�ݒ� */
        evtime = Timer              '�҂����߂̎���
        ntemp0 = isub
        mtemp0 = jsub
        otemp0 = ksub
        ntemp0 = T_keisu_cset(ntemp0, T_keisu(T_keisuCont(1) - 1))  'ntemp0
'        mtemp0 = T_keisu_cset(mtemp0, T_keisu(T_keisuCont(1) - 1))  'mtemp0   '2010.11.24�폜
'        otemp0 = T_keisu_cset(otemp0, T_keisu(T_keisuCont(1) - 1))  'otemp0   '2010.11.24�폜
        Do
          DoEvents
          fintime = Timer2func     ' 2009.8.17
'          fintime = Timer         '���ݎ���
          diTime = diffTime(fintime, evtime)
          If lsub <> 0 Then x1dt = diTime / lsub
          ndata = (ntemp0 - ntemp) * x1dt + ntemp
          mdata = (mtemp0 - mtemp) * x1dt + mtemp
          odata = (otemp0 - otemp) * x1dt + otemp
          TempSet 2, ndata
          TempSet 3, mdata
          TempSet 4, odata
          If diTime >= lsub Then Exit Do
        Loop
        ntemp = ntemp0
        mtemp = mtemp0
        otemp = otemp0
        TempSet 2, ntemp
        TempSet 3, mtemp
        TempSet 4, otemp
      Case "R"    '/* ��p */
        Select Case Int(isub)
        Case 0    '��p��@�n�e�e
          CoolOFF
        Case 1    '��p��@�n�m
          CoolON
        Case 2    '��p���@�n�m
          CoolON
        End Select
    End Select
jp0:
    If i < 24 Then
      i = i + 1
    Else
    End If
    If com = "B" Then Exit Do
  Loop
'/* ���`�v���Z�X�A���^�]�J�n */
'/* �f�[�^��ǂݎ�� */
'/* �u�U�[��炷 */
  'Label2(4).Caption = ""
'-----------------------------------------------------------------------------
st2:
'/* �^�C�g���̕\�� */
'/* �^�������̕\�� */
'/* ���W�l���̕\�� */
'/* �����p�y���ʒu�ύX�g�\�� */
'  Label2(5).Caption = Format(roz(0), "0.0000")     '/* �˓����`para�@�� */
  Label2(6).Caption = Format(roz(0), "0.0000") & Format(roz(1), "0.0")     '/* �˓����`para�@���� */
'------------------------------------------------------------------------------
'/* ���`�J�n */
'    i_s_do = -1   ' Do Loop �́@��           '  st: �́@�O�ֈړ� 2007.11.26
  Do        '--------- DO LOOP�@�@LOOP�@�Q�@�i�O����2�Ԗڂ̂koop�j�@�A�����`�̉񐔕����
    DoEvents
    I_s0 = i_s
    i_s = i_s + 1
    i_s_do = i_s_do + 1
    js = 0
    ist0 = -1
    ist1 = -1
    ie0 = 0
    ie1 = 0
    ie2 = 0
    ie3 = 0
    S_StartTime = Timer
    stTime = Timer
    sevTime = Timer
    diTimeSR = -9999.99                        ' ���x�ݒ�@�r�q�̏�����
    iSRcount = 1                               ' ���x�ݒ�@�r�q�̏�����
    For ii = 0 To 10
      ct_t(ii) = 0
    Next ii    ' ���x�ݒ�@�r�q�̏�����
'
    Label4(T_keisuCont(1) - 1).ForeColor = T_keisuCol!(3)    '  �����@�s���N(�|�C���^�[�j
    Label6(T_keisuCont(1) - 1).ForeColor = T_keisuCol!(3)    '  �����@�s���N(�|�C���^�[�j
    Label11(T_keisuCont(1) - 1).ForeColor = T_keisuCol!(3)    '  �����@�s���N(�|�C���^�[�j
    Label4(T_keisuCont(1) - 1).BorderStyle = 1  '  �g�t���ɂ���
    Label6(T_keisuCont(1) - 1).BorderStyle = 1  '  �g�t���ɂ���
    Label11(T_keisuCont(1) - 1).BorderStyle = 1  '  �g�t���ɂ���
'
    iz3 = Z3_HoseiCont(2)   ' Z�␳�@�����{����@ZNo.�@�@�@�f07.9.27�@�ǉ�
    z(iz3) = z(iz3) + Z3_Hosei(T_keisuCont(1) - 1) '  �hZ3"�̕␳�lset
'/*  ����t�@�C�����Ɓ@�ۉ���~�񐔁@�\��
  Label2(2).Caption = gcoxFlName + " -" + Format(iHoonStopNo, "0000")
'/* �J�E���^�ւ̏o�͂��� */

    If i_s <> 0 Then
      InitDat(11) = InitDat(11) + 1  '���`�J�E���^�g�E�^��
'      InitDtSave                   ' E  ���`���save�@02.8.25 s.f.
      Label2(13).Caption = Str(InitDat(11))
    End If
'/* ���`�g�̕\�� */�@�@�@-------�@��ʕ\���́@�ŏ�
ejs1:       ' ----- Loop 3  �ifor Loop �́@�O��)�@�@-----------------
  lGphNo0 = 0
  lGphNo = 0
  MoniGraph Me.Picture1, lGphNo0, lGphNo
'/* �w���̕\�� */
'/* �x���̕\�� */
'/* ���Đ�������јg�\�� */
    sdt = Format(Int(stime / 60), "0") & "��" & Format(Int(stime) Mod 60, "0") & "�b"
    Label2(8).Caption = sdt
    Label2(1).Caption = Format(ishu, "0")
    Label2(9).Caption = Format(T_keisuCont(1), "0")
    InitDat(10) = i_s               '���`�J�E���^
'
''    �������Ԑ���@�����A����̕\��       for no uchigawa he idou
     Label7(0).Caption = Format(DkatJ(0), "0.0")
     Label7(1).Caption = Format(DkatJ(1), "0.0")
    If (katCflag = True) Then
         Label7(0).BackColor = TKatBackCol(1)
         Label7(1).BackColor = TKatBackCol(1)
     Else
         Label7(0).BackColor = TKatBackCol(0)
         Label7(1).BackColor = TKatBackCol(0)
    End If
''
'/* �J�E���^�ւ̏o�̓_�E�� */
    'InitDat(11) = InitDat(11) - 1   '���`�J�E���^�g�E�^��
    'InitDtSave
    'Label2(13).Caption = Str(InitDat(11))
'/* �f�[�^�̎�荞�� */'
'    stTime = Timer            DO loop �J�n����ց@�ړ��@10/5
    evtime = Timer
    iflg = 1
    ied = 0
    ttime = Time
    mTime = Timer
'-----------------------------------------------------------------------------------
'----------------------------- For Loop i�@�@�擪
    imax = ptime * 60
    For i = 1 To imax      ' ----- Loop 4  FOR Loop -----�@ptime*60��@���
    '
start:           ' ----- Loop 5  START:  GOTO START: Loop -----
'
    'finTime = Timer    '2002.10.09 KYOCERA
      DoEvents               '2005.12.17 OverFlow �΍� s.f.  2006.3.3 ���� s.f.
      ituflg = 0            '�@�^�C���A�b�vflg�̃��Z�b�g10/5
'/* ���`���̃h���C�u*/�@�@�@�f�@ist0�@���@���݂̎��R�}���hNo.�@�@�@���ꂼ��̎��R�}���h�I�����ɃJ�E���gUP
        If ist0 > 0 Then
          If ic(ist0 - 1) = 10 Then      '  /* ic(ist0-1)=10 �I���̈Ӗ��@*/
            ist0 = ist0 - 1
          End If
        End If
          sdt3 = DispSegm(ist0)
          Label2(12).Caption = sdt3
        If ist0 <> ist1 Then             '�@�V�����ĊJ�n����
            gOrgFlg = False                '���_���A����=TRUE
            ist1 = ist0
            sevTime = Timer              '������Z�O�����g�J�n����
'
            If (ist0 > 0 And ist0 < 11) Then   '�@�J�n���Ԃ̕\���@�����������p
               diTime1 = diffTime(sevTime, stTime)          '2002.10.09 KYOCERA
               sdt = Format(ist0, "0") & "=" & Format(Int(diTime1 / 60), "0") & ":" & Format(Int(diTime1) Mod 60, "00")       '2002.10.09 KYOCERA
            End If
'
            Select Case ic(ist0)  '-------- �����䃂�[�h�ԍ�
            Case 0, 8   '-------------------- �ʒu����
              List1.Enabled = True
              List2.Enabled = True
              List3.Enabled = True
              ppos = "SC JikuStart 0 8"
              Ready_Wait    '
              CtlDisp     'outp(DIO_P+3,9); �T�[�{ON & ���x���S12
              s_drive z(ist0), vel(ist0)
            Case 1, 3, 7   '-------------------- ���x����  '2004.3.8 sf
              ppos = "SC JikuStart 1 3 7"
              List1.Enabled = False
              List2.Enabled = False
              List3.Enabled = False
              m_l = vel(ist0)
              'm_l = vel(ist0) / 100
              If m_l > 100 Then m_l = 100            '�@20120808�@50�@���@100�@��
              setcm1 z(ist0)
              Ready_Wait    '
              CtlVelo       'outp(DIO_P+3,5);  ���x����֐؂�ւ�
              'while((inp(XCN_DT1)&0x01)!=0);
'
'�@--- 2012.6.24 Z�m�F�@���łɓ��B���Ă���ꍇ�́@���̾����Ă�
          If r_z() >= z(ist0) Then
            ist0 = ist0 + 1             '
            Label2(6).Caption = "�ʒu pass PC " & Str(ist0)
          End If
'
              
              Do       ' �u�J�E���^�[��v�v��ԒE�o�p
                DoEvents
                sts = C870Sts(3)   'sts=1�̎��@���������u-1�v�@sts=0�̎��s���������u0�v
                If (sts And &H1) = 0 Then Exit Do   '�uPULSE �� COMPARE ����v��ԁv��loop
              Loop
              'Label2(11).Caption = Format(m_l, "0.000") 'printf("%7.3f",m_l);
            Case 2    '-------------------- �_�~�[
              ppos = "SC JikuStart 2"
              List1.Enabled = True
              List2.Enabled = True
              List3.Enabled = True
              Ready_Wait    '
              CtlDisp     'DioOut 12,1  �ʒu���� '  02.10.1 �ǉ�
              Ready_Wait    '
              ServoON     'outp(DIO_P+3,1);
            Case 9    '-------------------- �I��
              ppos = "SC JikuStart 9"
              List1.Enabled = True
              List2.Enabled = True
              List3.Enabled = True
              Ready_Wait    '
              CtlDisp     'outp(DIO_P+3,9);
              genten
              'Ready_Wait
              For ii = 1 To 180          '/* ����R�p�̏����� */
                r_z_dum(ii) = 0#
              Next ii
              i_ts = 1
              r_z_ave = 0#
            End Select
        End If
'
           fintime = Timer2func     ' 2009.8.17
'       fintime = Timer         '2002.10.09 KYOCERA   fintime:���ݎ���
'
'/* �^�C���A�b�v���� */
      '2002.10.09 KYOCERA
        If ist0 < 0 Then GoTo sj1:
'
'        For itu = 1 To 2000            ' 2010.1.16 �V�݁@LongTime���f�@20100130 for next 20 �� 500 '20103.6 500 -> 2000
          fintime = Timer2func        ' 2010.1.16 �V�݁@LongTime���f�ǉ��ɔ���
          tudiffTime = diffTime(fintime, sevTime)
          If ((ic(ist0) < 10) And (tudiffTime > (t0(ist0) * 1.2))) Then ' 2010.3.10 20100405  LongTime��t0(ist0)*1.2 �֕ύX, 20100407 tc(ist0)<10 �ǉ�--->ic(ist0)=10�́@�I���̈Ӗ�
             sdt = "��ѱ��� skip  " & Format(tudiffTime, "0.0")   ' 2010.3.10
             Label2(6).Caption = sdt     ' 2010.3.10
             GoTo TimeUpEnd:    '2010.3.6 �ύX�@for-next����߁Alongtime���傫��������timeup���[�`�����X�L�b�v
          End If
'          If tudiffTime < LongTime Then Exit For
'        Next itu
'
        If ((ic(ist0) < 10) And (tudiffTime > t0(ist0))) Then '2002.10.16 KYOCERA 2002.10.17 KYOCERA     '10/4
             ituflg = 1
             sdt = "��ѱ���" & Format(tudiffTime, "0.0")
             sdt = sdt & " " & Format(t0(ist0), "0.0") & " " & Format(ist0 + 1, "0")
             Label2(6).Caption = sdt
'
                RecEmgDtSave sdt3, sdt1, sdt2
                gemgmsg = "��ѱ���"
                hijyou              '����~����
                iFlg_hijyou = 1     '   �^�C���A�b�v
                GoTo eend:
'
'              ist0 = ist0 + 1             '/�^�C���A�b�v�Ŏ��̃X�e�b�v   '2002.10.16 KYOCERA
'            GoTo TimeUpEnd:
'             GoTo jscmdend:              '�@�I���M���������щz��    10/12 sf
        End If
TimeUpEnd:
'
'/* �I���M���̏��� */
        Select Case ic(ist0)
        Case 0, 8   '/* �ʒu����̏ꍇ */
          ppos = "SC JkE 0 8"
          If (C870Sts(1) And 1) = 0 Then
             ist0 = ist0 + 1
          End If
        Case 1    '/* ���x����̏ꍇ */
            ppos = "SC JkE1"
          pdt = pres(ist0)
          pp = p(ist0)
          pml = m_l
            ppos = "SC JkE1 -1cal"
'
          cal_pid pdt, pp, pml
            ppos = "SC JkE1 cal_pid"
          sts = C870Sts(3)  'status3 ��ǂ�
             ppos = "SC JkE1 sts=C870"
         If (sts And &H1) <> 0 Then      ' �����Łu-1�v�@�@�s�����Łu0�v
            ist0 = ist0 + 1             '/* �ʒu�B���ŏI�� */
            zzz = r_z()
            Label2(6).Caption = "�ʒu pass CNT " & Str(ist0) & " " & Str(zzz)   '11/2 sf & SP7  180901
'            Label2(6).Caption = "�ʒu pass CNT " & Str(ist0)   '11/2 sf
            rstcm1   '  compareter reset
'            Ready_Wait    '
         Else                       ' 2008.2.21  �ύX�@�P�b�ɂP��s���߂����m�F��
'
           If Int(mTime) = Int(Timer) Then
             If r_z() >= z(ist0) Then
               ist0 = ist0 + 1             '
               Label2(6).Caption = "�ʒu pass PC " & Str(ist0)
             End If
           End If
         End If
         ppos = "SC JkE1 r_z -1"
'''  Err.Raise 6  for test '''
        Case 3    '/* ���x����@�˓����`�̏ꍇ  2002.7.10 ls21_tc���R�s�[ */
           ppos = "SC JkE3"
          pdt = pres(ist0)
          pml = m_l
          pp = p(ist0)
           ppos = "SC JkE3 -1cal"
          cal_pid pdt, pp, pml
           ppos = "SC JkE3 cal_pid"
          sts = C870Sts(3)  'status3 ��ǂ�
           ppos = "SC JkE3 sts=C870"
          If (sts And &H1) <> 0 Then
            ist0 = ist0 + 1             '/* �ʒu�B���ŏI�� */
            zzz = r_z()
            Label2(6).Caption = "�ʒu pass CNT " & Str(ist0) & " " & Str(zzz)   '11/2 sf & SP7 180901
'            Label2(6).Caption = "�ʒu pass CNT " & Str(ist0)   '11/2 sf
            rstcm1   '  compareter reset
         Else                       ' 2008.2.21  �ύX�@�P�b�ɂP��s���߂����m�F��
           If Int(mTime) = Int(Timer) Then
             If r_z() >= z(ist0) Then
               ist0 = ist0 + 1             '
               Label2(6).Caption = "�ʒu pass PC " & Str(ist0)
             End If
           End If
         End If
'
          If r_z() < z(ist0) Then
'            r_z_now = r_z()                    '2008.2.23 �ړ�
              ppos = "SC JkE3 r_z -2"
            If Int(tsTime) <> Int(mTime) Then
              tsTime = mTime                  '/* �P�b�O�ƁA�Q�b�O�� */
              r_z_now = r_z()                    '2008.2.23 �ړ�
              If Abs(r_z_now - r_z_ave) < epsilon Then
                ist0 = ist0 + 1               '/* it_ts��A���@epsilon�ȉ� */
              Else                            '/* �Ł@�˓��B���ŏI�� */
                r_z_dum(i_ts) = r_z_now
                r_z_ave = 0#
                For ii = 1 To it_ts
                   r_z_ave = r_z_ave + r_z_dum(ii)
                Next ii
                r_z_ave = r_z_ave / it_ts
                i_ts = i_ts + 1
                If i_ts > it_ts Then i_ts = 1
              End If
            End If
          End If
        Case 7    '/* ���x����@�㎲�Փ˔���t�@�@�@�@�@�@�@�@�@2004.3.8 s.f. ������u�V�v�ǉ��@�@��������@*/
'�@�@�@  �@�@�@�@/*�@�w�舳�͂�荂�����͂��R�b�ȏ㑱���������~�@�@*/
          ppos = "SC JkE7"
          pdt = pres(ist0)
          pp = p(ist0)
          pml = m_l
          cal_pid pdt, pp, pml
          sts = C870Sts(3)  'status3 ��ǂ�
          If (sts And &H1) <> 0 Then
            ist0 = ist0 + 1             '/* �ʒu�B���ŏI�� */
             zzz = r_z()
            Label2(6).Caption = "�ʒu pass CNT " & Str(ist0) & " " & Str(zzz)   '11/2 sf & SP7  180901
'           Label2(6).Caption = "�ʒu pass CNT " & Str(ist0)   '11/2 sf
            rstcm1   '  compareter reset
'            Ready_Wait    '
'            Do                 'Do Loop  ' 2005.11.22 �폜�@�@��x�ǂ񂾂�status��reset�����B2�x�ǂݕs�I�I
'              DoEvents
''              sts = C870Sts(3)          'status3 ��ǂ�    10/4  sf
''              If (sts And &H1) <> 0 Then Exit Do           10/4 sf
'               If r_z() >= z(ist0) Then Exit Do             '10/4
'            Loop
          Else                       ' 2008.2.21  �ύX�@�P�b�ɂP��s���߂����m�F��
            If Int(mTime) = Int(Timer) Then        '�@�P�b��1��`�F�b�N
              If r_z() >= z(ist0) Then
                ist0 = ist0 + 1             '
                Label2(6).Caption = "�ʒu pass PC " & Str(ist0)
              End If
            End If
          End If
'
'
          If Int(tsTime) <> Int(mTime) Then '2008.2.23 �ύX 1�b��1��`�F�b�N
             tsTime = mTime                  '/* �P�b�O�Ɣ�r */
             bpre = r_pres()
             If iFlg_hijyou = 6 Then     '6=r_pres 1�g���z��
                gemgmsg = gemgmsg + " 1�g���z��"
                hijyou        '����~����
                GoTo eend:
             End If
'
             If bpre > pdt Then                ' 2008.2.18 �ύX
'               If bpre > pdt * 0.7 Then
               i_ts = i_ts + 1               '/* i_ts��A�����ā@���͂��w��l�ȏ� */
                If i_ts > 3 Then
                  gemgmsg = "������@�V"
                  hijyou        '����~����
                  'getch
                  iFlg_hijyou = 2    '    ������ 7�@error
                  GoTo eend:
                End If
             End If
          End If                                 '/*     '2004.3.8�@�����܂Ł@*/
       Case 9    '�I��
          ppos = "SC JkE9"
          sts = C870Sts(1)
          If (sts And 1) = 0 Then
            ist0 = ist0 + 1     '/* ���� */
            If Abs(r_z()) > 0.1 Then
              Label2(6).Caption = "���_�s��"
              ist0 = ist0 - 1
              genten              '���_�o��
            End If
          Else
            '/* �J�E���^�Ƀ[������������ */
            Ready_Wait
            Counter0
          End If
        End Select
''                                                  ' 2007.12.21 delete  ���x����l�̕\��
'      Select Case ic(ist0)                          ' 2004.3.12 s.f
'           Case 1, 3, 7                             ' 2005.11.4 s.f �폜
'                Label7(0).Caption = nout
'                Label7(1).Caption = v
''
'            Case 0, 2, 8, 9
''                �������ԏ�������̕\��
'                 Label7(0).Caption = Format(DkatJ(0), "0.0")
'                 Label7(1).Caption = Format(DkatJ(1), "0.0")
'      End Select
jscmdend:                               '������R�}���h�@������  10/4 sf
'
'''/* �G���[�\�� */     ' �A���[���\���@�P�b�ɂP��`�F�b�N�\���ց@090817 �ύX�@�i���ֈړ��j
''      If ArmChk <> 0 Then               '�A���[�����b�Z�[�W
''        frmerr_sign.Show   'ALM�o��
''      Else
''        Unload frmerr_sign
''      End If
'''
'/* �v���Z�X���s */
sj1:
      If iflg = 1 Then                          '�@iflg=1�@�O�̺���ޏI���̃t���O
        com = scom(js + flindex)                '�@js�@�́@�R�}���h��No.
        isub = sisub(js + flindex)
        jsub = sjsub(js + flindex)
        ksub = sksub(js + flindex)
        lsub = slsub(js + flindex)
        js = js + 1                              '�@js���@�����p�Ɂ@�P�i�߂Ă���
'
        evtime = Timer                  '  '05.12.17 evtime �J�E���g�J�n�������֐ݒu�@s.f.
'
        sdt = com & Right(Space(7) & Format(isub, "0"), 7)    ' ����ނ̕\��
'
        If ((Left(com, 1) = "S") Or (Left(com, 1) = "L")) Then
          sdt = sdt & Right(Space(7) & Format(jsub, "0"), 7)
          sdt = sdt & Right(Space(7) & Format(ksub, "0"), 7)
          sdt = sdt & Right(Space(7) & Format(lsub, "0"), 7)
        Else
          sdt = sdt
        End If
        Label2(7).Caption = sdt
      End If
        '�V�X�e�����f�B/* ����~�̏ꍇ�͐��`���~ */
          sts1 = SystemReadyChk()   '�V�X�e�����f�B or ����~
          sts2 = AutoChk()          '������ԁH
          If sts1 = 0 Or sts2 = 0 Then
            gemgmsg = ArmEmgMsgChk$()
            iFlg_hijyou = 10              '����~ү���ނ̂�������
            FrmMenuFlg = False              '���j���[���甲����Ƃ�false
            NextView = 1
            Exit Do                         '�@Loop�����яo��������
          End If
        '
          Select Case Left(com, 1)
          Case "D"    '------------ ���`���̌^�̗L��   ��,�\�A�A�\���@��111�@�S���^����@100���S�ȏ�Ȃ琬�`���^����
             ppos = "SC Proc D"
             If (isub = 0) Then     '�ݔۃZ���T�[�`�F�b�N
               If (KataChk() > 3) Then                '  2004.10.30  �^�ݔۃ`�F�b�N�p�Z���T�̓���m�F�p
                 sdt = "DC�@�ݔۃZ���T�[�ُ�i�^�L��I�I�j"
                 Label2(6).Caption = sdt
'
                  sdt2 = sdt2 & sdt
                  RecEmgDtSave sdt3, sdt1, sdt2
                  gemgmsg = "DC �^�L��"
                  hijyou        '����~����
                  iFlg_hijyou = 3          '�@DC�@error�@�^�L��
                  GoTo eend:
               Else
                  GoTo scend:
               End If
            End If                                 '  2004.10.30  �^�ݔۃ`�F�b�N�p�Z���T�̓���m�F�p
'
            If (KataChk() < 4) Or (Karauchiflg = True) Then '���`���Ɍ^�������@�@'08.4.22
              fintime = Timer2func     ' 2009.8.17
'               fintime = Timer       ' ���ݎ��ԁ@�@�@�@'2006.3.3�@�@�ǉ��@s.f.
              If (diffTime(fintime, evtime) < isub) Then
                 iflg = 0             ' ���Ԗ��B�̏ꍇ
              Else
                 idmy = js            '�@���ԑ҂��I���̏ꍇ�@�@js�@=�@���̃R�}���h��No.�@�@(�ŏ��ɓǂݎ�邽�߁A�l��1�i��ł���j
                 Do
                   DoEvents
                   dmy = scom(idmy + flindex)          '�@���̃R�}���h��ǂݎ��
                   If "LA" = dmy Then  '----- �R�}���hLA�܂Ői�߂�
                     js = idmy          '�@�@LA������������@���̃R�}���hNo.���@LA�́@No.�ɃZ�b�g
                     '------------- LA�����������玟�ɁA�Z�O�����g�����[�h�W�܂Łi9�̂Q�O�܂Łj�i�߂�
                     Do
                       DoEvents
                       If ic(ist0) = 8 Then
                         ist0 = ist0 - 1
                         sevTime = Timer        '  2005.12.17 Timeup�h�~ �O�̂��� s.f.
                         Exit Do
                       End If
                       ist0 = ist0 + 1
                       If ist0 > 50 Then   '�G���[
'
                         sdt = "DC����� ist0 > 50 �װ"
                         Label2(6).Caption = sdt
                         RecEmgDtSave sdt3, sdt1, sdt2
                         gemgmsg = "DC�@�G���[�@4"
                         hijyou        '����~����
                         iFlg_hijyou = 4        '�@DC�@�R�}���h�G���[
                         GoTo eend:
'
                       End If
                     Loop
                   '
                     Exit Do
                   End If
                   idmy = idmy + 1
                   If idmy > 50 Or "EN" = dmy Then '�G���[
'
                         sdt = "DC����� ist0 > 50 �װ"
                         Label2(6).Caption = sdt
                         RecEmgDtSave sdt3, sdt1, sdt2
                         gemgmsg = "DC�@�G���[�@5"
                         hijyou        '����~����
                         iFlg_hijyou = 5          '�@�@DC�R�}���h�G���[
                         GoTo eend:
'
                   End If
                 Loop
'
                 iflg = 1                    '�@����ޏI������
                 idcflg(1) = 1               '  DC�t���O�@�^��=1�@�^�L=0
'                 evtime = Timer              ' 2005.12.17�@s.f.
                  sevTime = Timer             ' 2005.12.17 �O�̂���
              End If
            Else
              idcflg(1) = 0             '  �^������ꍇ�@idcflg=0�ɂ��Ĕ�����
            End If                    '�@�^������ꍇ�͂��̂܂ܔ�����
'
          Case "L"    '------------ ���`���Ɍ^�������������̔�ѐ�Ԓn
             ppos = "SC Proc L"
             If (KataChk() < 4) Then GoTo caselend: '�^������
             If (iflghoonStop = False) And (iflg5Stop = False) Then GoTo caselend:
'                      ------------  �^������A���@�ۉ���~�t���O�@ON�̎��̏���
'             DoEvents           '2005.12.17  OverFlow �΍� s.f.
             iflg = 0
             Command2(0).Enabled = False
             Command2(9).Enabled = False
'�@�@�@�@�@�@�@�@�@�@�@------------
              ntemp0 = isub
              mtemp0 = jsub
              otemp0 = ksub
              ntemp0 = T_keisu_cset(ntemp0, T_keisu(T_keisuCont(1) - 1))  'ntemp0
              mtemp0 = T_keisu_cset(mtemp0, T_keisu(T_keisuCont(1) - 1))  'mtemp0�@�@'2010.11.24  �폜 2012.1.5 �ۉ���~�����ŕ���
              otemp0 = T_keisu_cset(otemp0, T_keisu(T_keisuCont(1) - 1))  'otemp0�@�@'2010.11.24  �폜�@2012.1.5 �ۉ���~�����ŕ���
              TempSet 2, ntemp0
              TempSet 3, mtemp0
              TempSet 4, otemp0
'
''              DoEvents           '2005.12.17  OverFlow �΍� s.f.
              If (iflghoonStop = True) Then
                 Label12(0).Visible = True
                 Label12(1).Visible = True
                 Label12(2).Visible = True
                 Command1.Visible = True
                 Label12(0).Caption = "�ۉ���~��"
                 Label12(1).Caption = " �o�ߎ���"
                  
         ''  �@�ۉ���~�@���ԑ҂��@-----------------------------
                 hs5_sttime = Timer
                 imachi = 60 * 60 - 1          '  �҂����ԁ@60������
                 Do
                   DoEvents
                   hs5_fintime = Timer
                   hs5_difft = diffTime(hs5_fintime, hs5_sttime)
                   If (hs5_difft < imachi) And (iHoteikanryou = 0) Then
                      If (Int(hs5_difft) <> Int(hs5_diffTold)) Then
                          Label12(2).Caption = Format(Int(hs5_difft / 60), "  00��") + Format(Int(hs5_difft) Mod 60, " 00�b")
                          hs5_diffTold = hs5_difft
                      End If
                       Else
                          Exit Do              '�@���ԑ҂��I��
                       End If
                 Loop
'
                 Label12(0).Visible = False
                 Label12(1).Visible = False
                 Label12(2).Visible = False
                 Command1.Visible = False
                  iHoteikanryou = 1
                  iflg = 1
                  GoTo caselend2:
'                 iflghsmsg = MsgBox("�ۉ�����~�@���������܂����H", 48, "�ۉ�����~��")  '��~�������͑҂�
'              DoEvents         '2005.12.17  OverFlow �΍� s.f.  2006.5.18 �ǉ�
              End If
'
              If (iflg5Stop = True) Then
                 Label12(0).Visible = True
                 Label12(1).Visible = True
                 Label12(2).Visible = True
                 Label12(0).Caption = "5����~��"
                 Label12(1).Caption = " �ĊJ�܂� "
'
         ''  �@5���ԕۉ���~�@���ԑ҂��@-----------------------------
                 hs5_sttime = Timer
                 imachi = 5 * 60 - 1          '  �҂����ԁ@�T������
                 Do
                   DoEvents
                   hs5_fintime = Timer
                   hs5_difft = diffTime(hs5_fintime, hs5_sttime)
                   If (hs5_difft < imachi) Then
                      If (Int(hs5_difft) <> Int(hs5_diffTold)) Then
                          Label12(2).Caption = Format(Int((imachi - hs5_difft) / 60), "  0��") + Format(Int((imachi - hs5_difft)) Mod 60, " 0�b")
                          hs5_diffTold = hs5_difft
                          End If
                       Else
                          Exit Do              '�@���ԑ҂��I��
                       End If
                 Loop
'
                 Label12(0).Visible = False
                 Label12(1).Visible = False
                 Label12(2).Visible = False
              End If
  '
'�@�@�@�@�@�@�@�@�@�@�@-------------�@�I���̏���
caselend2:    TempSet 2, ntemp    ' ���̉��x�ɖ߂��ďI��
              TempSet 3, mtemp
              TempSet 4, otemp
'
             If (iflghoonStop = True) Then
                  iHoonStopNo = iHoonStopNo + 100  ' �ۉ���~�񐔂̃J�E���g�A�b�v
                  iflghoonStop = False   ' �t���O�����Z�b�g
                  Command2(9).BackColor = CmndColoff(9)    '�R�}���h�{�^���̐F��߂�
              End If
              If (iflg5Stop = True) Then
                iHoonStopNo = iHoonStopNo + 1  ' �ۉ���~�񐔂̃J�E���g�A�b�v
                iflg5Stop = False   ' �t���O�����Z�b�g
                Command2(0).BackColor = CmndColoff(0)    '�R�}���h�{�^���̐F��߂�
              End If
              
             Command2(0).Enabled = True
             Command2(9).Enabled = True

'
              sevTime = Timer     '�@������R�}���h���^�C���A�b�v���Ȃ��悤�Ɂ@sevtime�̃��Z�b�g
              evtime = Timer      '  2005.12.17  �O�̂���  s.f.
'
caselend:   iHoteikanryou = 1
            iflg = 1            '����𔲂���ƏI��
'              evtime = Timer             ' 2005.12.17�@s.f.
'
          Case "H"    ' �����\�[�N�@�@�@�h�g�b�h
             ppos = "SC Proc H"
             fintime = Timer2func     ' 2009.8.17
'             fintime = Timer      ' ���ݎ��ԁ@�@�@'�@2006.3.3�@�ǉ��@s.f.
             If (lSokuFlg = True And diffTime(fintime, evtime) < isub) Then
               iflg = 0
             Else
               iflg = 1
               lSokuFlg = False
               Command2(8).BackColor = SokuCor(0)
'               evtime = Timer             ' 2005.12.17�@s.f.
             End If
'
          Case "S"    '/* �`�s�b���x�ݒ� */
             ppos = "SC Proc S"
            If Mid(com, 2, 1) = "R" Then             ' SR�̏ꍇ  ���F�֘A�������@Do�@Loop�@Top�ɂ���
               fintime = Timer2func     ' 2009.8.17
'               fintime = Timer
               diTime = diffTime(fintime, stTime)    ' 0.1�b�ɂP�񉷓x��荞�݁i�T����{�j
               If ((diTime - diTimeSR) > 0.1) Then
                   AdRead dt(), adFlg   'AD�{�[�h����@���x�Ǎ�
                   ct_dummy = dt(0) '   '���x�Ǎ��@�@�P�F���`���@IH�q�[�^�[
                   ct_dummy = T_keisu_cread(ct_dummy, T_keisu(T_keisuCont(1) - 1))
                   ct_t(0) = ct_t(0) + ct_dummy
'
                   ct_dummy = dt(5) '   '���x�Ǎ��@�@�U�F���`���@��^
                   ct_dummy = T_keisu_cread(ct_dummy, T_keisu(T_keisuCont(1) - 1))
                   ct_t(5) = ct_t(5) + ct_dummy
'
                   ct_dummy = dt(6) '   '���x�Ǎ��@�@�V�F���`���@���^
                   ct_dummy = T_keisu_cread(ct_dummy, T_keisu(T_keisuCont(1) - 1))
                   ct_t(6) = ct_t(6) + ct_dummy
'
                   iSRcount = iSRcount + 1
                   diTimeSR = diTime
                   iflg = 0
                   If iSRcount > 5 Then
                      ct_t(0) = ct_t(0) / 5
                      ct_t(5) = ct_t(5) / 5
                      ct_t(6) = ct_t(6) / 5
                      ntemp0 = isub
                      ntemp0 = T_keisu_cset(ntemp0, T_keisu(T_keisuCont(1) - 1)) 'ntemp0
                      mtemp0 = jsub
'                      mtemp0 = T_keisu_cset(mtemp0, T_keisu(T_keisuCont(1) - 1)) 'mtemp0   '2010.11.24 �폜
                      otemp0 = ksub                                                         '2010.11.24 jsub -> ksub
'                      otemp0 = T_keisu_cset(otemp0, T_keisu(T_keisuCont(1) - 1)) 'otemp0   '2010.11.24 �폜
                      ntemp0 = ct_t(0) + ntemp0
                      mtemp0 = ct_t(5) + mtemp0
                      otemp0 = ct_t(6) + otemp0
                      ntemp = ntemp0
                      mtemp = mtemp0
                      otemp = otemp0
                      TempSet 2, ntemp
                      TempSet 3, mtemp
                      TempSet 4, otemp
                      ct_t(0) = 0: ct_t(5) = 0: ct_t(6) = 0
                      Label2(6).Caption = "SR= " & Format(Int(ntemp), "000") & Format(Int(mtemp), "  000") & Format(Int(otemp), "  000")
                      iSRcount = 1
                      iflg = 1
'                      evtime = Timer             ' 2005.12.17�@s.f.
                   End If
               End If
            Else
             ppos = "SC Proc SA"
              fintime = Timer2func     ' 2009.8.17
'              fintime = Timer
              diTime = diffTime(fintime, evtime)        'SA�̏ꍇ
'              DoEvents     '2005.12.17  OverFlow �΍� s.f.  2006.5.18 �ǉ� �폜
             ppos = "SC Proc SA af dev"
              If lsub <> 0 Then x1dt = diTime / lsub
              ntemp0 = isub
              mtemp0 = jsub
              otemp0 = ksub
              ntemp0 = T_keisu_cset(ntemp0, T_keisu(T_keisuCont(1) - 1))  'ntemp0
'              mtemp0 = T_keisu_cset(mtemp0, T_keisu(T_keisuCont(1) - 1))  'mtemp0    ' 2010.11.24 �폜
'              otemp0 = T_keisu_cset(otemp0, T_keisu(T_keisuCont(1) - 1))  'otemp0    ' 2010.11.24 �폜
              ndata = (ntemp0 - ntemp) * x1dt + ntemp
              mdata = (mtemp0 - mtemp) * x1dt + mtemp
              odata = (otemp0 - otemp) * x1dt + otemp
              TempSet 2, ndata
              TempSet 3, mdata
              TempSet 4, odata
              If diTime >= lsub Then
                iflg = 1
                ntemp = ntemp0
                mtemp = mtemp0
                otemp = otemp0
                TempSet 2, ntemp
                TempSet 3, mtemp
                TempSet 4, otemp
'                evtime = Timer             ' 2005.12.17�@s.f.
              Else
                iflg = 0
              End If
            End If
          Case "P"    '/* �ړ�������̋쓮 */
             ppos = "SC Proc P"
            If Mid(com, 2, 1) = "W" Then
              Beep
              ist0 = ist0 + 1
              sevTime = Timer          '2005.12.17�@�O�̂��߁@s.f.
'              evtime = Timer          '2002.10.09 KYOCERA               ' 2005.12.17�@s.f.
            End If
            If Mid(com, 2, 1) = "R" Then
              iflg = 0
              If ist0 <> ist1 Then iflg = 1
              If isub = 4 And ist0 = 0 Then iflg = 1
'              If iflg = 1 Then evtime = Timer             '2002.10.09 KYOCERA               ' 2005.12.17�@s.f.
              If iflg = 1 Then sevTime = Timer             '2005.12.17�@s.f.
             End If
          'evTime = Timer
          Case "K"    '/* ���M */
             ppos = "SC Proc K"
            Select Case isub
            Case 1
              HeatON
            Case 0
              HeatOFF
            End Select
          Case "N"
             ppos = "SC Proc N"
            If Mid(com, 2, 1) = "S" Then
              If isub = 1 Then hdt = hdt
              If isub = 0 Then hdt = hdt
            End If
          Case "W"    '/* ����p */
             ppos = "SC Proc WC"
            Select Case isub
            Case 1
              SuireiON
            Case 0
              SuireiOFF
            End Select
          Case "R"    '/* �K�X��p */
            If (Mid(com, 2, 1) = "C") Then
                 ppos = "SC Proc R"
                Select Case isub
                Case 2
                 CoolON
              Case 1
                CoolON
              Case 0
                CoolOFF
              End Select
            Else
                ppos = "SC Proc RM"
                Select Case isub
                Case 1
                    SuireiON
                Case 0
                    SuireiOFF
                End Select
            End If
          Case "T"    '/* �`�s�b�P�̉��x�̓ǂݎ�� */
             ppos = "SC Proc T"
            sdata = TempRdMold(0)    '�X���[�u���x
            sdata = T_keisu_cread(sdata, T_keisu(T_keisuCont(1) - 1)) 'ndata
            If (Mid(com, 2, 1) = "L" And sdata > isub) Or (Mid(com, 2, 1) = "G" And sdata < isub) Or (Mid(com, 2, 1) = "E" And (sdata > (isub + 20) Or sdata < (isub - 20))) Then
'            If (Mid(com, 2, 1) = "L" And sdata > isub) Or (Mid(com, 2, 1) = "G" And sdata < isub) Then
              iflg = 0
            Else
              If iflg = 2 Then iflg = 1 Else iflg = 2
'              evtime = Timer             ' 2005.12.17�@s.f.
            End If
          Case "J"    '/* ���ԑ҂� */
             ppos = "SC Proc J"
            DoEvents             ' 2006.5.18  �ǉ��@s.f
            fintime = Timer2func     ' 2009.8.17
'            fintime = Timer      ' ���ݎ��ԁ@�@�@�@�f2006.3.3�@�ǉ��@s.f.
            diTime1 = diffTime(fintime, stTime)
            diTime2 = diffTime(fintime, evtime)
            If (Mid(com, 2, 1) = "S" And diTime1 >= isub) Or (Mid(com, 2, 1) = "C" And diTime2 >= isub) Then
              iflg = 1
'              evtime = Timer             ' 2005.12.17�@s.f.
            Else
              iflg = 0
            End If
          Case "C"
             ppos = "SC Proc C"
            Select Case Mid(com, 2, 1)
            Case "P"    '���`�I���ʒu�@�`�F�b�N
              cp_z = r_z()
              Label5(0).Caption = " cp=   " & Format(cp_z, "0.000")
            Case "C"    '�@���ԃ`�F�b�N
              If isub > 3 Then
                  ict = 5
              Else
                ict = isub + 2
              End If
              fintime = Timer2func     ' 2009.8.17
'              fintime = Timer         '���ݎ���
              cc_time(isub) = diffTime(fintime, stTime)
              sdt = " cc" & Format(isub, "0") & "= " & Format(Int(cc_time(isub) / 60), "0") & ":" & Format(Int(cc_time(isub)) Mod 60, "00")        '2002.10.09 KYOCERA
              Label5(ict).Caption = sdt
              If isub = 3 Then
                diTime1 = diffTime(cc_time(isub), cc_time(isub - 1))
                katJ = diTime1
                sdt = " cc3-2= " & Format(Int(diTime1 + 0.5), "0") & "s"
                Label5(6).Caption = sdt
              End If
'
          Case "T"    '�@���x�`�F�b�N
            If isub > 2 Then
                ict = 2
              Else
                ict = isub
            End If
            ct_temp(isub - 1) = TempRdMold(0) '�X���[�u���x 300��-2000��
            ct_temp(isub - 1) = T_keisu_cread(ct_temp(isub - 1), T_keisu(T_keisuCont(1) - 1))
            sdt = " ct" & Format(isub, "0") & "=   " & Format(ct_temp(isub - 1), "0.0") & "��"
            Label5(ict).Caption = sdt
          End Select
          Case "X"    '�����I���M���i���`�J�n�j
             ppos = "SC Proc X"
            Select Case Mid(com, 2, 1)
              Case "R"    '���`�J�n [�����I���܂ő҂�]
            '
                TrnsReqON  '�����˗��M��Ch21�o�� (�����I������)
                'WaitSec 1.5  '
            '
                Do
              '-------------- �s���j�v�ǂ�
 '                 LS21S_Monitor    '2006.12.21 �폜 s.f
                  'DioInput 13, sts        '�����I���H
                  sts = TrnsFinChk()      '�����I���H
                  If sts = 1 Then
                    TrnsReqOFF            '�����˗��M���n�e�e
                    Exit Do
                  End If
                  DoEvents           '  ���Ӂ@����DoEvents���@Do�@����Ɉڂ��Ɓ@�듮�삷��B�@�����I��2��҂��ɂȂ�I�I
                Loop
'
'               --- �^�@No.�̕\���@��񑗂�@---
                kataNoPnt = kataNoPnt + 1
                If kataNoPnt > katamax Then kataNoPnt = 0
'
                For iii = katamax To 0 Step -1
                    Label13(iii).Caption = kataNoHyj(katamax - iii + kataNoPnt + katamax + 1 + Val(kataNo(10)))
                Next iii
'
                If (i_s_do) < katamax - 1 Then
                    For iii = kataNoPnt + 1 To katamax
                        Label13(iii).Caption = "��"
                    Next iii
                End If
'
' ---           �^�m���D�@�P�񑗂芮��
              Case "W"    '���`�I��
              End Select
          Case "E"    '/* �I���@���{�b�g���� */
             ppos = "SC Proc E"
             DoEvents
            If iflg <> 99 Then
              iflg = 0
              If r_z() > 2 Then
                genten
                'Ready_Wait    'while((inp(AX_STS)&1)!=0);
              End If
              TrnsReqON       '�����˗��M��Ch21�o��
              WaitSec 1.5     '
              '�����\���M��Ch15��҂�
              'DioInput 15, sts
              'If sts = 1 Then
                iflg = 99
              'End If
              isp = 0
            Else
             'DioInput 13, sts    '�����I���M��Ch13��҂�
              sts = TrnsFinChk()      '�����I���H
              If sts = 1 Then
                TrnsReqOFF        '�����˗��M��OFF
                GoTo send:
              Else
              End If
            End If
scend:
          End Select
cjump:
'
  '-------------- �s���j�v�ǂ�
'          LS21S_Monitor�@�@�@�@�@2005.6.4�@�폜s.f.
'
'          DoEvents
          lEmgFlg = SystemReadyChk()  '����~�̊m�F
          If Int(mTime) = Int(Timer) And lEmgFlg <> 0 Then GoTo start:
           mTime = Timer
'
'                   Loop 5�@ start: ����@�����܂Ł@�����Ƀ��[�v
' ---------------- /* 1�b��1�񉺂ɔ����� ��ʕ\���o��*/  ------------------------
'
          ppos = "SC 1sec Disp 1"
'           /* ���́@�o�h�c����@�o���P�T�@�Ȃ瑬�x�@�[�� */
          If ist0 >= 0 Then
            If p(ist0) > 15 Then
              DaVoltOut 1, 0        ' 0V D/A ch=1
            End If
          End If
'/* �G���[�\�� */                       ' 09.8.17 �ォ�炱���ֈ����z��
      If ArmChk <> 0 Then               '�A���[�����b�Z�[�W
        frmerr_sign.Show   'ALM�o��
      Else
        Unload frmerr_sign
      End If
'
    KeikaTime(i) = it + 1
'/*�@���x��荞�� */
'          DoEvents               '2005.12.17 OverFlow �΍� s.f.
          atemp(i, 0) = TempRdMold(0)   '�X���[�u���x 0V-300�� 1V-1300��
          atemp(i, 0) = T_keisu_cread(atemp(i, 0), T_keisu(T_keisuCont(1) - 1))
          atemp(i, 1) = TempRdMold(5)                 '�ヂ�[���h���x
          atemp(i, 1) = T_keisu_cread(atemp(i, 1), T_keisu(T_keisuCont(1) - 1))
          atemp(i, 2) = TempRdMold(6)                 '�����[���h���x
          atemp(i, 2) = T_keisu_cread(atemp(i, 2), T_keisu(T_keisuCont(1) - 1))
'
'* ���`���ʒu�̎�荞�� */
          ppos = "SC 1sec Disp 2"
          aposi(i) = r_z()
'/* �^���͂̎�荞�� */
          ppos = "SC 1sec Disp 3"
          apre(i) = r_pres()
          If iFlg_hijyou = 6 Then     '6=r_pres 1�g���z��
             gemgmsg = gemgmsg + " 1�g���z��"
             hijyou        '����~����
             GoTo eend:
          End If
'
'/* ���x���z�̕\�� */
'/* �^�����̃v���b�g */
'/* ���W�l�̃v���b�g */
          lGphNo = i
          GphDataSet lGphNo0, lGphNo
          MoniGraph Me.Picture1, lGphNo0, lGphNo
          lGphNo0 = lGphNo
jo0:
'/* �e��f�[�^�̉�ʉ��\�� �P�@*/
          DoEvents           '2006.5.18 OverFlow �΍� s.f. �ǉ�
          sdt1 = Format(atemp(i, 0), "  0.0��     ")
          sdt1 = sdt1 & Format(apre(i), "0.00kgf    ")
          sdt1 = sdt1 & Format(aposi(i), "0.000mm   ")
          Label2(14).Caption = sdt1
'/* �e��f�[�^�̉�ʉ��\�� �Q */
          it0 = Timer                                                          ' 10/5
          it = diffTime(it0, stTime)
          sdt2 = Format(Int(it / 60), "  0��")
          sdt2 = sdt2 & Format(Int(it) Mod 60, " 0�b")      '2002.10.09 KYOCERA
          sdt2 = sdt2 & "     ct " & Format(diffTime(it0, evtime), "0.0")
          sdt2 = sdt2 & "     st " & Format(diffTime(it0, sevTime), "0.0")
'          sdt2 = sdt2 & "tt   " & Format(diffTime(it0, stTime), "0.0")    '2005.11.23 ���ԍ팸�̂��ߍ폜
          Label2(11).Caption = sdt2
'
'/* �����\�� */
          Label8.Caption = Time$
'
'/* ��ޯĈʒu�ύX�@*/
          'If FrmMenuFlg = False Then GoTo eend:
      Next i   '----- Loop 4  -- For Loop�@i�@�I�[�@ 1��̐��`�܂���1�񕪂̉�ʕ\���I��
      js = js - 1        '�@js=�@���̃R�}���h�̔ԍ��@�@�i1�߂��Ă���j
      GoTo ejs1:      '�@Loop 3�@---/* �\���I���Ō���ʂ� */�i���񕪁@��ʕ\���ցj
'
'
' ----------------  1�񕪂̐��`�I���@--------------------------------------
send:
'    ---- /* �^�N�g�^�C���̎Z�o�@*/ ----
      ppos = "SC 1��end"
      iSeikeiTorF_flg = True
      iSento_flg = 0            '�擪�_�~�[�׸ރ��Z�b�g
'�@�@�@�@�@�@�@�@�@�@�@�@�@���`��@����̐��`�̗L�����m�F�i���`�񐔗p�j�@�@'100405�@if�̒����炱���ֈړ�
        idcflg(3) = idcflg(2)          '  idcflg(3) �P��O
        idcflg(2) = idcflg(1)          '  idcflg(2) ����
'
      If i_s > 0 Then       ' ���`�P��ڂ́@i_s=0�@�Ł@Pass�B�@�@'100306�@�폜�B'100405 �����@else�ȍ~�ǉ��@"���`����|�C���^�[����" �o�O�C���̂��߁@s.f.
        If idcflg(2) = 1 Then
           i_s = i_s - 1                  '��̎��́@���`�񐔁|�P �����V���b�g
           InitDat(11) = InitDat(11) - 1  '���`�J�E���^�g�E�^���̖߂�
           iSeikeiTorF_flg = False
        Else
          If idcflg(3) = 1 Then
            i_s = i_s - 1                 '�_�~�[�̎��́A�����V���b�g�@�i����^�L��{1��O����@���@�_�~�[�@�j
            InitDat(11) = InitDat(11) - 1  '���`�J�E���^�g�E�^���̖߂�
            iSeikeiTorF_flg = False
            iSento_flg = 1                ' �擪�_�~�[�׸�
          End If
        End If
      Else                                '���`����@i_s=0�@�̎��@�ʏ����@�@'100405�ǉ�
        If idcflg(2) = 1 Then
           i_s = i_s - 1                  '��̎��́@���`�񐔁|�P �����V���b�g
           iSeikeiTorF_flg = False
        End If
      End If          '100306�@�폜�B�@���`����@�|�C���^�[����o�O�C���̂��߁@s.f.  '100405 �����ielse�����ǉ��j
      If i_s = 0 Then iSeikeiTorF_flg = False
'
'     stime = i
      endTime = Timer
      stime = diffTime(endTime, stTime)         '  10/5
      InitDtSave            '�@�f�[�^save�@�i���`�񐔁j
'
'
' --- �������Ԃ̕��ϒl�v�Z�@�@���݂̌^No��T_keisuCont(1)-1�@�A�@���݂���@�S���O�܂ł̕��ϒl
'     --- ���񂪁@�_�~�[�@�̏ꍇ�A�@�����f�[�^(KatJ)�����Z�b�g�i0�ցj
      If iflgKataTorF(T_keisuCont(1) - 1) = False Then
        For ikat = 0 To 3
          kaatsuJ(T_keisuCont(1) - 1, ikat) = 0#
        Next ikat
      End If
'�@�@----�@�f�@�^�ύX���̎�舵�� �^���s�ςŐV�K�^�ɓ���ւ��i�O�Ƀ��Z�b�g����j
     If (i_s > 0) And (i_s <> I_s0) Then    '   -----------------�������Ԑ��䃋�[�`���@start
                                            '  --------- �L���Ȑ��`���ǂ����̔���
                
'
        kaatsuJ(T_keisuCont(1) - 1, 0) = katJ    '  katJ=����̉�������
' ---                                            ' �������ԕ��ϒl�@����̉������ԁ@�d�݁i�E�F�C�g�j2.0�ց@�@2007.11.21
        avekatJ(T_keisuCont(1) - 1) = (kaatsuJ(T_keisuCont(1) - 1, 0) * 2 + kaatsuJ(T_keisuCont(1) - 1, 1) + kaatsuJ(T_keisuCont(1) - 1, 2) + kaatsuJ(T_keisuCont(1) - 1, 3)) / (4 + 1)
'
        kjdisp = Format(InitDat(11), "000") & "  "
        kjdisp = kjdisp & Format(T_keisuCont(1), "00") & "  "
        kjdisp = kjdisp & Format(avekatJ(T_keisuCont(1) - 1), "000") & "  "
        For ikat = 0 To 3
           kjdisp = kjdisp & Format(kaatsuJ(T_keisuCont(1) - 1, ikat), "000") & "  "
        Next ikat
'     --- �VT�W���v�Z ---�@�@���ϒl�ƍ���̉������ԂŁ@�]��
'       ---�@�i�P�j���ϒl���@����������ɂ��邩�H
        If ((avekatJ(T_keisuCont(1) - 1)) > DkatJ(1)) Then
              T_keisu_dum = T_keisu(T_keisuCont(1) - 1) + 0.001      '������傫���ꍇ�@+0.001          DkatJ(1)=����l
        Else
             If (avekatJ(T_keisuCont(1) - 1) >= DkatJ(0)) Then
                  T_keisu_dum = T_keisu(T_keisuCont(1) - 1)       ' ����ȉ��A���A�����ȏ�Ȃ�@���̒l�̂܂�
             Else
                  T_keisu_dum = T_keisu(T_keisuCont(1) - 1) - 0.001  '������菬�����ꍇ�@-0.001      DkatJ(�O)=�����l
             End If
        End If
'
'       ---�@�i�Q�j����̉������Ԃ��@����������ɂ��邩�H
        If ((katJ <= DkatJ(1)) And (katJ >= DkatJ(0))) Then
              T_keisu_dum = T_keisu(T_keisuCont(1) - 1)             ''����̉������Ԃ��@����Ɖ��������Ȃ�@T�W���́@�ς��Ȃ��I
        End If
'       ---�@�i3�j����̉������Ԃ��@�����ȉ����H
        If (katJ < DkatJ(0)) Then
              T_keisu_dum = T_keisu(T_keisuCont(1) - 1) - 0.001           ''����̉������Ԃ��@����Ɖ��������Ȃ�@T�W���́@�ς��Ȃ��I
        End If
'     --- �\�� ---
        kjdisp = kjdisp & Format(T_keisu_dum, "0.000") & "  " & Format(T_keisu(T_keisuCont(1) - 1), "0.000") & "  "
        List2.AddItem kjdisp, 0
'     ---'����v�Z�p�@�f�[�^�X�V ----
        For ikat = 3 To 0 Step -1
          kaatsuJ(T_keisuCont(1) - 1, ikat + 1) = kaatsuJ(T_keisuCont(1) - 1, ikat)
        Next ikat
      End If                ' ---------------------- �������Ԑ��䃋�[�`���@end
'
'     --- �������Ԏ�������@���{/pass�@---
      katDflag = True        '  ---  "0" ���@�����Ă��Ȃ����m�F������
      For ikat = 0 To 3
        If (kaatsuJ(T_keisuCont(1) - 1, ikat) < 1) Then katDflag = False
      Next ikat
'�@�@�@�@�@---�@��������@���{�ۊm�F
      If ((katCflag = True) And (katDflag = True) And (iflgKataTorF(T_keisuCont(1) - 1) = True) And (iSeikeiTorF_flg = True)) Then T_keisu(T_keisuCont(1) - 1) = T_keisu_dum
'      If ((katCflag = True) And (kaatsuJ(T_keisuCont(1) - 1, 3) <> 0) And (iflgKataTorF(T_keisuCont(1) - 1) = True)) Then T_keisu(T_keisuCont(1) - 1) = T_keisu_dum
'
      Label4(T_keisuCont(1) - 1).Caption = Format(T_keisu(T_keisuCont(1) - 1), "0.000")
'
'     --- �������ԁA�b���l�@�`�k�`�q�l�\�� ---
        AlmON = False
        Almdisp = Format(ishu, "0") & "-" & Format(T_keisuCont(1), "0")
        If (katJ < AkatJ(0)) Or (katJ > AkatJ(1)) Then
            AlmON = True
            Almdisp = Almdisp & " k= " & Format(katJ, "0")
        End If
        If (cp_z < Acp(0)) Or (cp_z > Acp(1)) Then
            AlmON = True
            Almdisp = Almdisp & " C= " & Format(cp_z, "0.000")
        End If
        If iSeikeiTorF_flg = False Then AlmON = False
        If iflgKataTorF(T_keisuCont(1) - 1) = False Then AlmON = False
        If i_s < 1 Then AlmON = False
        If AlmON = True Then List3.AddItem Almdisp, 0
'�@ --- /*�@���`�f�[�^�̕\���i���X�g�\���j�@*/  2002.12.3 sf  ---
'        InitDat(11)=���`�񐔁i�V���b�g���j
'
      Rec_of_Mold = Format(InitDat(11), "000") & "  " & Format(ishu, "0") & " " & Format(T_keisuCont(1), "0") & " "
      Rec_of_Mold = Rec_of_Mold & " " & Format(z(iz3), "000.00") & "    "
      Rec_of_Mold = Rec_of_Mold & " " & Format(Int(ct_temp(0)), "000") & "�� " & Format(Int(ct_temp(1)), "000") & "��  "
      Rec_of_Mold = Rec_of_Mold & " " & Format(Int(cc_time(1) / 60), "0") & ":" & Format(Int(cc_time(1)) Mod 60, "00") & " "
      Rec_of_Mold = Rec_of_Mold & " " & Format(Int(cc_time(2) / 60), "0") & ":" & Format(Int(cc_time(2)) Mod 60, "00") & " "
      Rec_of_Mold = Rec_of_Mold & " " & Format(Int(cc_time(3) / 60), "0") & ":" & Format(Int(cc_time(3)) Mod 60, "00") & " "
      diTime1 = diffTime(cc_time(3), cc_time(2))
      Rec_of_Mold = Rec_of_Mold & "  " & Format(Int(diTime1 + 0.5), "000") & "s "
      Rec_of_Mold = Rec_of_Mold & "  " & Format(cp_z, "000.000") & " "
      Rec_of_Mold = Rec_of_Mold & "  " & Format(Int(stime / 60), "0") & ":" & Format(Int(stime) Mod 60, "00") & " "
      Rec_of_Mold = Rec_of_Mold & "  " & Format(T_keisu(T_keisuCont(1) - 1), "0.000") & "    " & Format(Z3_Hosei(T_keisuCont(1) - 1), "0.000")
      Rec_of_Mold = Rec_of_Mold & "  " & Format(avekatJ(T_keisuCont(1) - 1), "000") & "  " & Format(iHoonStopNo, "###0") & "  " & Format(zzz, "000.000")  'SP7 180901
'      Rec_of_Mold = Rec_of_Mold & "  " & Format(avekatJ(T_keisuCont(1) - 1), "000") & "  " & Format(iHoonStopNo, "###0")
      If AlmON = True Then Rec_of_Mold = Rec_of_Mold & "  " & Almdisp
      List1.AddItem Rec_of_Mold, 0    ' �h�A0�h�@�ǉ��@2004.8.18
'
      RecDtSave Rec_of_Mold
'
'
'' /* ���x�W���A�����␳�f�[�^�̃J�E���g�A�b�v
      Label4(T_keisuCont(1) - 1).ForeColor = T_keisuCol!(2)  '  �����F�����ɖ߂�
      Label6(T_keisuCont(1) - 1).ForeColor = T_keisuCol!(2)  '  �����F�����ɖ߂�
      Label11(T_keisuCont(1) - 1).ForeColor = T_keisuCol!(2)  '  �����F�����ɖ߂�
      Label4(T_keisuCont(1) - 1).BorderStyle = 0  '  �g�Ȃ��ɖ߂�
      Label6(T_keisuCont(1) - 1).BorderStyle = 0  '  �g�Ȃ��ɖ߂�
      Label11(T_keisuCont(1) - 1).BorderStyle = 0  '  �g�Ȃ��ɖ߂�
'     *** Z3�̒l���@�߂�
          z(iz3) = z(iz3) - Z3_Hosei(T_keisuCont(1) - 1) '  �hZ3"�̕␳�lreset
'     *** �|�C���^�[�J�E���g�A�b�v
      If (i_s > 0) And (i_s <> I_s0) Then
        T_keisuCont(1) = T_keisuCont(1) + 1       ' �|�C���^�[�̃J�E���g�A�b�v
      End If
      If T_keisuCont(1) > (T_keisuCont(0)) Then     ' 1���̏I��� count up
        T_keisuCont(1) = 1
        ishu = ishu + 1
      End If
'
      T_keisuCont(2) = T_keisuCont(1)           ' ** �|�C���^�[��Buckup **
      T_keisuCont(3) = T_keisuCont(0)           ' ** �^���@��Buckup **
      ishu_bkup = ishu                          ' ** ?�T�ځ@�́@Backup **
'       --- Saikaiflg �@���@false�@��
      Saikaiflg = False
'/* �f�[�^�̕ۑ��@*/
      If lDtSaveFlg = True Then
        ResDtSave i_s, stime
        lDtSaveFlg = False          '�f�[�^�Z�[�u�@��t����
'
        Command2(5).BackColor = CmndColoff(1)    ' off gray
        Command2(5).Caption = "Save"
       End If
'
' ScreenCopy iflgSCopy=1 or 2  �̏ꍇ�AScreenCopy
    Select Case iflgSCopy
        Case 1
                If (iSeikeiTorF_flg) Or (iSento_flg = 1) Then ' ��������or�擪�̏ꍇ�@copy
                Call SaveWindowPic(True, False)     'Active Window�̕ۑ�
                iflgSCopy = 0          'ScreenCopy�@��t����
                Command2(2).BackColor = CmndColoff(0)
                End If
        Case 2
                If (iSeikeiTorF_flg) Then    ' �L�����Ď��̂݁��擪��а�̎��̖{�^���@copy
                Call SaveWindowPic(True, False)     'Active Window�̕ۑ�
                iflgSCopy = 0          'ScreenCopy�@��t����
                Command2(2).BackColor = CmndColoff(0)
                End If
    End Select
''    If (iSeikeiTorF_flg) Or (iSento_flg = 1) = True Then    ' 20130425 �����V���b�g���f���폜
''         If iflgSCopy = True Then
''             Call SaveWindowPic(True, False)     'Active Window�̕ۑ�
''         End If
''         iflgSCopy = 0          'ScreenCopy�@��t����
''         Command2(2).BackColor = CmndColoff(0)
''   End If
'''
 '/* �G�f�B�Ƃ�������Ă�����@�G�f�B�b�g */
      If FrmMenuFlg = False Then Exit Do            '�I����������Ă���ƃ��j���[���甲����Ƃ�false
      If EditFlg = True Then '�G�f�B�^�N��
         ied = 1             '�G�f�B�^�N���́@doLoop�̊O�Ŏ��{�@06.3.3 sf
         Exit Do
      End If
'/* ������~��Ԃł���Β�~ */
      sts1 = SystemReadyChk()   '�V�X�e�����f�B or ����~
      sts2 = AutoChk()          '������ԁH
      If sts1 = 0 Or sts2 = 0 Then    '1��ڊm�F
        For idum = 1 To 10000: iidum = iidum + 1: Next idum   'Delay
        sts1 = SystemReadyChk()   '�V�X�e�����f�B or ����~
        sts2 = AutoChk()          '������ԁH
        If sts1 = 0 Or sts2 = 0 Then    '�V�X�e�����f�B or ����~�́@�Q��ڊm�F
          gemgmsg = ArmEmgMsgChk$()
          iFlg_hijyou = 10            '����~���̏��Z�[�u
          FrmEmg.Show 1               '�@����~�\��
          FrmMenuFlg = False              '���j���[���甲����Ƃ�false
          NextView = 1
          SeikeiOFF         '����~���̏��u '���`OFF�@�ҋ@��
          HeatOFF          '����~���̏��u
          CoolOFF          '����~���̏��u
          ServoOFF         '����~���̏��u
        Exit Do
        End If
      End If
  Loop    '-------------------- DO LOOP�@Loop 2�@�i�O����2�Ԗڂ̃��[�v�j
'/*�@���������̂Ƃ��́@do�@Loop���甲����@�ύX�@060303 s.f
'/*  �G�f�B�b�g��������Ă����� �@ied=1�@*/
  If ied = 1 Then '�G�f�B�^�N��
      Command2(3).BackColor = CmndColoff(3)  '�F��߂�
      EditFlg = False      '�G�f�B�^�N������
      MYEdit.Show 1
      ied = 0
      c = 0
      GoTo st:             '/* �G�f�B�b�g���[�h�ł���΁@�����ɃW�����v */
'     --------------    '  Loop�@�P�@�@�i�ŊO���[�v�j ---------------
  End If
'/* �G�f�B�b�g���[�h�ł���΁@�����ɃW�����v */
'  If ied <> 0 Then GoTo st:
'
'   �����łȂ���ΏI����
'/* �\�����M���[���ɂ��A�n�e�e���� */
eend:
  If iFlg_hijyou > 0 Then              '����~���痈����
    RecEmgDtSave sdt3$, sdt1$, sdt2$ & gemgmsg
  End If
  SeikeiOFF          '���`OFF�@�ҋ@��
  HeatOFF
  CoolOFF
  ServoOFF
'/* cox�f�[�^�̂g�c�ւ̏����o�� */
'    ����I����  ���۰��ް���save
      coxDtSet
      coxDtSave gcoxFldir & gcoxFlName
''
''  ���`�f�[�^�t�@�C���ց@�R���g���[���f�[�^��ǉ��@�@2009.9.12�ǉ�
      RecDtSave999
''
  If FrmMenuFlg = False Then             '���j���[���甲����Ƃ�false
    FrmMenuFlg = True                    '���j���[���甲����Ƃ�false
    Select Case NextView
    Case 1
      Unload Me
      PGM_Menu.Show
    Case 2 '���`�i�V���O���j
      NQD70_SC.Show
    Case 3  '���`�i�_�u���j
    Case 4  'I O �`�F�b�N
      IOChk.Show
    Case 5  '�X�P�[���ύX
      LS21_GphScale.Show
    Case 6  '�ǂݏo��
    Case 7  '������
    Case 8  'edit
      Unload Me
      MYEdit.Show
    Case Else
      Unload Me
      PGM_Menu.Show
    End Select
  End If
  If iFlg_hijyou = 0 Then Unload Me       '����~���痈�����́A�����Ȃ��i��ʎc���j
  PGM_Menu.Show
'
Exit Sub
'
errHandler:
  SeikeiOFF          '���`OFF�@�ҋ@��
  HeatOFF
  ServoOFF
  CoolOFF
'
  RecEmgDtSave sdt3, sdt1, sdt2
  If Err.Number <> 0 Then
     sdt1 = "�G���[�ԍ��F" & Err.Number
     sdt2 = "��ۼު�Ė��F" & Err.Source & "  " & ppos
     sdt3 = "�G���[���e�F" & Err.Description
  End If
  RecEmgDtSave sdt1, sdt2, sdt3
  gemgmsg = Err.Number & Err.Description
  hijyou        '����~����
'
End Sub
Private Sub genten()
'--------------
  C870Genten
  gOrgFlg = True                       '���_���A����=TRUE
  OrgON
  gOrgStartFlg = True   '2002.10.17 KYOCERA
End Sub

Private Sub GphXSet()
Dim i%
  For i = 0 To ptime * 60 + 10
    TPass(i) = i
  Next i
End Sub

Private Sub GphDataSet(i0%, i1%)
Dim i%
  For i = i0 To i1
    Templ(i) = atemp(i, 0)
    Templu(i) = atemp(i, 1)   '��^���x
    Templd(i) = atemp(i, 2)   '���^���x
    Press(i) = apre(i)
    ZAxis(i) = aposi(i)
  Next i
End Sub

Private Function DispSegm$(ist0%)
Dim sdt$
    If ist0 < 0 Then Exit Function
    sdt = Right(Space(2) & Format(ist0, "0"), 2)
    sdt = sdt & Right(Space(4) & Format(seg_num(ist0), "0"), 4)
    sdt = sdt & Right(Space(4) & Format(ic(ist0), "0"), 4)
    sdt = sdt & Right(Space(12) & Format(z(ist0), "0.000"), 12)
    sdt = sdt & Right(Space(7) & Format(vel(ist0), "0.0"), 7)
    sdt = sdt & Right(Space(6) & Format(pres(ist0), "0"), 6)
    sdt = sdt & Right(Space(7) & Format(t0(ist0), "0.0"), 7)
    sdt = sdt & Right(Space(7) & Format(p(ist0), "0.0"), 7)
    DispSegm = sdt
'    Label2(12).Caption = sdt
End Function
Private Function EmgChk%()
Dim ch%, sts%
  ch = 1
  DioInput ch, sts
  If sts = 0 Then
    EmgChk = True
  Else
    EmgChk = False
  End If
End Function

Private Sub Timer2_Timer()
    If r_z > 0.1 Then
        OrgOFF
    Else
        OrgON
    End If
    
    'Label6(0).Caption = "���_ = " & gOrgIL
    'Label6(1).Caption = r_z
End Sub

'�X�N���[���̃X�i�b�v�V���b�g���N���b�v�{�[�h�ɕۑ��y�ш���@�{�́@�@�@�@�@�i273�j '

Private Sub SaveWindowPic(Optional ActWind As Boolean = True, _
                                    Optional PrintOn As Boolean = False)
'�X�N���[���̃X�i�b�v�V���b�g���N���b�v�{�[�h�ɕۑ��y�ш���@�@�@�@�@�@�@�@�@�i273�j '
'�t�H�[����Command�{�^�����Q�\��t���Ă����ĉ������B
'�@ Option Explicit�@�@ 'SampleNo=273�@WindowsXP VB6.0(SP5) 2003.03.30
'�L�[�X�g���[�N���V�~�����[�g����(P1065)

    Dim MyFileName As String, PicData As Picture, OsVer As Single
    Dim sngSt As Single
'
    Clipboard.Clear
    OsVer = CreateObject("SysInfo.SYSINFO").OSVersion

    If ActWind Then
    '�A�N�e�B�u �E�B���h�E�̃X�i�b�v�V���b�g���擾����
    '�ȉ��̂Q���@�ǂ�ł�OK(Win98SE/WinXP/Win95�j
    '�ǂ̕��@�ł���L�m�F�@��͓������삵�܂��̂�MS�̃T���v���̕��@���g�p
        Call keybd_event(VK_LMENU, &H56, _
                                KEYEVENTF_EXTENDEDKEY Or 0, 0)
        Call keybd_event(VK_SNAPSHOT, &H79, _
                                KEYEVENTF_EXTENDEDKEY Or 0, 0)
        Call keybd_event(VK_SNAPSHOT, &H79, _
                                KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
        Call keybd_event(VK_LMENU, &H56, _
                                KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
'�@�@�@�@==================== ������ł������悤�ł� ==================
'�@�@�@�@Call keybd_event(VK_LMENU, 0, _
�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@KEYEVENTF_EXTENDEDKEY Or 0, 0)
'�@�@�@�@Call keybd_event(VK_SNAPSHOT, 0, _
�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@KEYEVENTF_EXTENDEDKEY Or 0, 0)
'�@�@�@�@Call keybd_event(VK_SNAPSHOT, 0, _
�@�@�@�@�@�@�@�@�@�@�@KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
'�@�@�@�@Call keybd_event(VK_LMENU, 0, _
�@�@�@�@�@�@�@�@�@�@�@KEYEVENTF_EXTENDEDKEY Or KEYEVENTF_KEYUP, 0)
    ElseIf ActWind = False And OsVer < 5 Then
    '��ʑS�̂̃X�i�b�v�V���b�g���擾����(Win98SE/Win95)
        Call keybd_event(VK_SNAPSHOT, 1, KEYEVENTF_EXTENDEDKEY, 0)
        Call keybd_event(VK_SNAPSHOT, 1, KEYEVENTF_EXTENDEDKEY Or _
                                                                          KEYEVENTF_KEYUP, 0)
    Else
    '��ʑS�̂̃X�i�b�v�V���b�g���擾����(WinXP)
        Call keybd_event(VK_SNAPSHOT, 0, KEYEVENTF_EXTENDEDKEY, 0)
        Call keybd_event(VK_SNAPSHOT, 0, KEYEVENTF_EXTENDEDKEY Or _
                                                                          KEYEVENTF_KEYUP, 0)
    End If
'
    sngSt = Timer                           ' Windows7 �ɂ́A���̒x��Loop���K�v
    Do While Timer - sngSt < 0.5
       DoEvents
    Loop
'
    '�N���b�v�{�[�h���Ƀr�b�g�}�b�v�`���̃f�[�^�����邩���ׂ�
    If Clipboard.GetFormat(vbCFBitmap) Then
        '�t�@�C��������������
        MyFileName = App.path & "\..\data\" & gcoxFlName$ & Format$(Now, "yymmddhhmmss") & ".BMP"
        '�\���f�[�^�[���r�b�g�}�b�v�`���̃f�[�^�ŕۑ�
        Set PicData = Clipboard.GetData
        Call SavePicture(PicData, MyFileName)
        If PrintOn Then
            '�������ꍇ
            With Printer
                .ScaleMode = vbMillimeters
                .PaperSize = vbPRPSA4
                .Orientation = vbPRORLandscape
                .PaintPicture PicData, 10, 0
                .EndDoc
            End With
        End If
    Else
        MsgBox "�ۑ��o���܂���ł����B"
    End If
End Sub
'
'
'
'Private Sub Command1_Click()
''�A�N�e�B�u�E�C���h�E�݂̂��N���b�v�{�[�h�ɃR�s�[
'    Call SaveWindowPic(True, False)     '�������ꍇ�́@True �ɐݒ�
'End Sub
'
'Private Sub Command2_Click()
''�X�N���[���S�̂��N���b�v�{�[�h�ɃR�s�[
'    Call SaveWindowPic(False, False)
'End Sub


