VERSION 5.00
Begin VB.Form LS21_TC 
   BackColor       =   &H00C0C0C0&
   Caption         =   "1�񐬌`"
   ClientHeight    =   8532
   ClientLeft      =   132
   ClientTop       =   420
   ClientWidth     =   11856
   ControlBox      =   0   'False
   FillColor       =   &H00C0C0C0&
   BeginProperty Font 
      Name            =   "�l�r �o�S�V�b�N"
      Size            =   12
      Charset         =   128
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00800000&
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8532
   ScaleWidth      =   11856
   StartUpPosition =   3  'Windows �̊���l
   Begin VB.Timer Timer2 
      Interval        =   200
      Left            =   720
      Top             =   3240
   End
   Begin VB.CommandButton Command2 
      Caption         =   "�����\�[�N"
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
      Height          =   240
      Index           =   8
      Left            =   5880
      MaskColor       =   &H8000000F&
      Style           =   1  '���̨���
      TabIndex        =   62
      Top             =   655
      Visible         =   0   'False
      Width           =   1308
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���`�J�n(�w��)"
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
      Index           =   7
      Left            =   1920
      TabIndex        =   61
      Top             =   480
      Width           =   1692
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  '�ׯ�
      Height          =   372
      Index           =   0
      Left            =   3720
      TabIndex        =   60
      Text            =   "4"
      Top             =   480
      Width           =   492
   End
   Begin VB.CommandButton Command2 
      Caption         =   "���`�J�n(3��)"
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
      Index           =   6
      Left            =   1920
      TabIndex        =   59
      Top             =   120
      Width           =   1692
   End
   Begin VB.CommandButton Command2 
      Caption         =   "S"
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
      Index           =   5
      Left            =   5040
      TabIndex        =   58
      Top             =   480
      Visible         =   0   'False
      Width           =   468
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
      Left            =   120
      TabIndex        =   57
      Top             =   840
      Visible         =   0   'False
      Width           =   1692
   End
   Begin VB.CommandButton Command2 
      Caption         =   "V �G�f�B�^���"
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
      TabIndex        =   55
      Top             =   480
      Width           =   1668
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   200
      Left            =   720
      Top             =   2160
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  '�ׯ�
      AutoRedraw      =   -1  'True
      BackColor       =   &H00800000&
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   9
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   5500
      Left            =   1800
      ScaleHeight     =   5472
      ScaleWidth      =   8376
      TabIndex        =   9
      Top             =   1860
      Width           =   8400
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '�_��
         Index           =   7
         X1              =   6696
         X2              =   6696
         Y1              =   0
         Y2              =   5436
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '�_��
         Index           =   6
         X1              =   5040
         X2              =   5040
         Y1              =   0
         Y2              =   5436
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '�_��
         Index           =   5
         X1              =   3348
         X2              =   3348
         Y1              =   0
         Y2              =   5436
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '�_��
         Index           =   4
         X1              =   1656
         X2              =   1656
         Y1              =   0
         Y2              =   5436
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '�_��
         Index           =   3
         X1              =   0
         X2              =   8352
         Y1              =   1080
         Y2              =   1080
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '�_��
         Index           =   2
         X1              =   0
         X2              =   8352
         Y1              =   2196
         Y2              =   2196
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '�_��
         Index           =   1
         X1              =   0
         X2              =   8352
         Y1              =   3312
         Y2              =   3312
      End
      Begin VB.Line Line2 
         BorderColor     =   &H00C0C0C0&
         BorderStyle     =   3  '�_��
         Index           =   0
         X1              =   0
         X2              =   8352
         Y1              =   4392
         Y2              =   4392
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
      TabIndex        =   6
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label7 
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
      Height          =   204
      Index           =   1
      Left            =   11100
      TabIndex        =   95
      Top             =   4560
      Width           =   612
   End
   Begin VB.Label Label7 
      Caption         =   "  T�W��"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   0
      Left            =   10320
      TabIndex        =   94
      Top             =   4560
      Width           =   612
   End
   Begin VB.Label Label10 
      BackColor       =   &H80000005&
      Caption         =   "Label10"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   12
         Charset         =   128
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   252
      Left            =   10400
      TabIndex        =   83
      Top             =   660
      Width           =   1284
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   9
      Left            =   10320
      TabIndex        =   93
      Top             =   7140
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   8
      Left            =   10320
      TabIndex        =   92
      Top             =   6876
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   7
      Left            =   10320
      TabIndex        =   91
      Top             =   6624
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   6
      Left            =   10320
      TabIndex        =   90
      Top             =   6360
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   5
      Left            =   10320
      TabIndex        =   89
      Top             =   6096
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   4
      Left            =   10320
      TabIndex        =   88
      Top             =   5844
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   3
      Left            =   10320
      TabIndex        =   87
      Top             =   5580
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   2
      Left            =   10320
      TabIndex        =   86
      Top             =   5316
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   1
      Left            =   10320
      TabIndex        =   85
      Top             =   5064
      Width           =   600
   End
   Begin VB.Label Label4 
      Caption         =   "Label4"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   0
      Left            =   10320
      TabIndex        =   84
      Top             =   4800
      Width           =   600
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
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
      Left            =   11160
      TabIndex        =   82
      Top             =   7440
      Width           =   612
   End
   Begin VB.Label Label8 
      Caption         =   "Label8"
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
      Index           =   0
      Left            =   10320
      TabIndex        =   81
      Top             =   7440
      Width           =   612
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  '����
      Caption         =   "cc3-2"
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
      Index           =   6
      Left            =   10320
      TabIndex        =   80
      Top             =   3700
      Width           =   1452
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   9
      Left            =   11100
      TabIndex        =   79
      Top             =   7140
      Width           =   660
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   8
      Left            =   11100
      TabIndex        =   78
      Top             =   6876
      Width           =   660
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   7
      Left            =   11100
      TabIndex        =   77
      Top             =   6624
      Width           =   660
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   6
      Left            =   11100
      TabIndex        =   76
      Top             =   6360
      Width           =   660
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   5
      Left            =   11100
      TabIndex        =   75
      Top             =   6096
      Width           =   660
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   4
      Left            =   11100
      TabIndex        =   74
      Top             =   5844
      Width           =   660
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   3
      Left            =   11100
      TabIndex        =   73
      Top             =   5580
      Width           =   660
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   2
      Left            =   11100
      TabIndex        =   72
      Top             =   5316
      Width           =   660
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   1
      Left            =   11100
      TabIndex        =   71
      Top             =   5064
      Width           =   660
   End
   Begin VB.Label Label6 
      Caption         =   "Label6"
      BeginProperty Font 
         Name            =   "�l�r �o�S�V�b�N"
         Size            =   7.8
         Charset         =   128
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   204
      Index           =   0
      Left            =   11100
      TabIndex        =   70
      Top             =   4800
      Width           =   660
   End
   Begin VB.Label Label5 
      Caption         =   "cc3"
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
      Index           =   5
      Left            =   10320
      TabIndex        =   69
      Top             =   3320
      Width           =   1452
   End
   Begin VB.Label Label5 
      Caption         =   "cc2"
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
      Index           =   4
      Left            =   10320
      TabIndex        =   68
      Top             =   2960
      Width           =   1452
   End
   Begin VB.Label Label5 
      Caption         =   "cc1"
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
      Index           =   3
      Left            =   10320
      TabIndex        =   67
      Top             =   2600
      Width           =   1452
   End
   Begin VB.Label Label5 
      Caption         =   "ct2"
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
      Left            =   10320
      TabIndex        =   66
      Top             =   2200
      Width           =   1452
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  '����
      Caption         =   "ct1"
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
      Left            =   10320
      TabIndex        =   65
      Top             =   1860
      Width           =   1452
   End
   Begin VB.Label Label5 
      BorderStyle     =   1  '����
      Caption         =   "cp1"
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
      Index           =   0
      Left            =   10320
      TabIndex        =   64
      Top             =   4120
      Width           =   1452
   End
   Begin VB.Label Label2 
      Alignment       =   1  '�E����
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
      TabIndex        =   63
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
      Left            =   10750
      TabIndex        =   56
      Top             =   80
      Width           =   950
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
      TabIndex        =   54
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
      TabIndex        =   53
      Top             =   8160
      Width           =   5136
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
      Index           =   11
      Left            =   6720
      TabIndex        =   52
      Top             =   8160
      Width           =   4980
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
      Index           =   10
      Left            =   8760
      TabIndex        =   51
      Top             =   960
      Width           =   2952
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�V���b�g���F"
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
      Left            =   8280
      TabIndex        =   50
      Top             =   120
      Width           =   1548
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�T�C�N���^�C���F"
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
      Left            =   7800
      TabIndex        =   49
      Top             =   396
      Width           =   2052
   End
   Begin VB.Label Label2 
      Alignment       =   1  '�E����
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
      Index           =   9
      Left            =   9720
      TabIndex        =   48
      Top             =   80
      Width           =   950
   End
   Begin VB.Label Label2 
      Alignment       =   1  '�E����
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
      Index           =   8
      Left            =   9720
      TabIndex        =   47
      Top             =   384
      Width           =   1980
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
      Left            =   1428
      TabIndex        =   46
      Top             =   7800
      Width           =   5136
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
      Width           =   1290
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
      Index           =   5
      Left            =   6720
      TabIndex        =   44
      Top             =   1560
      Width           =   5052
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
      Index           =   4
      Left            =   3240
      TabIndex        =   43
      Top             =   1560
      Width           =   3312
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "���`��ԁF"
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
      Index           =   1
      Left            =   1920
      TabIndex        =   42
      Top             =   1560
      Width           =   1296
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
      Height          =   210
      Index           =   31
      Left            =   9360
      TabIndex        =   41
      Top             =   7560
      Width           =   465
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
      Height          =   210
      Index           =   30
      Left            =   7275
      TabIndex        =   40
      Top             =   7560
      Width           =   870
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
      TabIndex        =   39
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
      Index           =   27
      Left            =   6660
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
      Index           =   26
      Left            =   4965
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
      Index           =   25
      Left            =   3270
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
      Height          =   216
      Index           =   24
      Left            =   1680
      TabIndex        =   34
      Top             =   7488
      Width           =   372
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
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   23
      Left            =   1230
      TabIndex        =   33
      Top             =   1260
      Width           =   660
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
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   22
      Left            =   1290
      TabIndex        =   32
      Top             =   1515
      Width           =   465
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      Index           =   20
      X1              =   1620
      X2              =   1764
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      Index           =   19
      X1              =   1620
      X2              =   1764
      Y1              =   2970
      Y2              =   2970
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      Index           =   18
      X1              =   1620
      X2              =   1764
      Y1              =   4050
      Y2              =   4050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      Index           =   17
      X1              =   1620
      X2              =   1764
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      Index           =   16
      X1              =   1620
      X2              =   1764
      Y1              =   6270
      Y2              =   6270
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FF00&
      Index           =   15
      X1              =   1620
      X2              =   1764
      Y1              =   7350
      Y2              =   7350
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00C00000&
      Index           =   14
      X1              =   1770
      X2              =   1770
      Y1              =   1875
      Y2              =   7375
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
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   21
      Left            =   1170
      TabIndex        =   31
      Top             =   1770
      Width           =   495
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
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   20
      Left            =   1290
      TabIndex        =   30
      Top             =   2880
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
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   19
      Left            =   1290
      TabIndex        =   29
      Top             =   3960
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
      ForeColor       =   &H0000FF00&
      Height          =   216
      Index           =   18
      Left            =   1320
      TabIndex        =   28
      Top             =   5076
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
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   17
      Left            =   1290
      TabIndex        =   27
      Top             =   6150
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
      ForeColor       =   &H0000FF00&
      Height          =   210
      Index           =   16
      Left            =   1290
      TabIndex        =   26
      Top             =   7230
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
      Height          =   210
      Index           =   15
      Left            =   540
      TabIndex        =   25
      Top             =   1260
      Width           =   660
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
      Height          =   216
      Index           =   14
      Left            =   600
      TabIndex        =   24
      Top             =   1512
      Width           =   492
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   13
      X1              =   1005
      X2              =   1149
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   12
      X1              =   1005
      X2              =   1149
      Y1              =   2970
      Y2              =   2970
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   11
      X1              =   1005
      X2              =   1149
      Y1              =   4050
      Y2              =   4050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   10
      X1              =   1005
      X2              =   1149
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H000000C0&
      Index           =   9
      X1              =   1005
      X2              =   1149
      Y1              =   6270
      Y2              =   6270
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
      X1              =   1155
      X2              =   1155
      Y1              =   1860
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
      TabIndex        =   23
      Top             =   1770
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
      TabIndex        =   22
      Top             =   2880
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
      TabIndex        =   21
      Top             =   3960
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
      TabIndex        =   20
      Top             =   5070
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
      TabIndex        =   19
      Top             =   6150
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
      TabIndex        =   18
      Top             =   7230
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
      Height          =   210
      Index           =   7
      Left            =   30
      TabIndex        =   17
      Top             =   1260
      Width           =   450
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
      Height          =   210
      Index           =   6
      Left            =   30
      TabIndex        =   16
      Top             =   1515
      Width           =   495
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   6
      X1              =   390
      X2              =   534
      Y1              =   1860
      Y2              =   1860
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   5
      X1              =   390
      X2              =   534
      Y1              =   2970
      Y2              =   2970
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   4
      X1              =   390
      X2              =   534
      Y1              =   4050
      Y2              =   4050
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   3
      X1              =   390
      X2              =   534
      Y1              =   5160
      Y2              =   5160
   End
   Begin VB.Line Line1 
      BorderColor     =   &H0000FFFF&
      Index           =   2
      X1              =   390
      X2              =   534
      Y1              =   6270
      Y2              =   6270
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
      Y1              =   1860
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
      TabIndex        =   15
      Top             =   1770
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
      TabIndex        =   14
      Top             =   2880
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
      TabIndex        =   13
      Top             =   3960
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
      TabIndex        =   12
      Top             =   5070
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
      TabIndex        =   11
      Top             =   6150
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
      TabIndex        =   10
      Top             =   7230
      Width           =   375
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "�R�����g�F"
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
      Index           =   0
      Left            =   1908
      TabIndex        =   8
      Top             =   1248
      Width           =   1272
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
      Left            =   3216
      TabIndex        =   7
      Top             =   1248
      Width           =   8508
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
      Index           =   2
      Left            =   3960
      TabIndex        =   5
      Top             =   936
      Width           =   4572
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  '����
      Caption         =   "����t�@�C�����F"
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
      Index           =   7
      Left            =   1944
      TabIndex        =   4
      Top             =   924
      Width           =   2028
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
      Left            =   7368
      TabIndex        =   3
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
      Index           =   1
      Left            =   5748
      TabIndex        =   2
      Top             =   360
      Visible         =   0   'False
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
      Left            =   5748
      TabIndex        =   1
      Top             =   72
      Width           =   1500
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
      Left            =   4488
      TabIndex        =   0
      Top             =   108
      Width           =   1272
   End
End
Attribute VB_Name = "LS21_TC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    LS21_TC
'
'            update: 2002.6.28 s.f  private sub cal_pid�@�폜
'            update: 2002.6.29 s.f "DC" ��������
'                                  "HC" �V�K�ǉ�
'�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@difftime�@��������
'            update: 2002.8.10 s.f roz(0),roz(1)��˓����`�����Ұ���
'            update: 2002.8.18 s.f DC���@���`�񐔂̖߂��@�ǉ��@(i_s=i_s-1)
'            update: 2002.8.29 s.f cp,ct,cc�f�[�^�\��'
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
'            �@�@�@�@�@�@�@�@�@�@�@�@cc3-cc2�\���@�ǉ�
'                                   SR�@�̏����ύX�@0.1�b�ɂP������ݸ�
'
'            update: 2002.11.1 s.f iPltMax �����l�@10�@->�@8�@�֕ύX
'            update: 2002.12.4 s.f ���`�f�[�^��save
'            update: 2003.07.10 HND �A���[���\�����́@���`�v���O�������s
'                                  frmerr_sign, FbiDio, LS21_SC
'            update: 2003.07.19  s.f.  1�񐬌`�̉񐔁@�U�|���R�֕ύX�@�P���@�̂�
'            update: 2003.09.11  s.f.  Plt1Jyun()�ց@WaitSec 1.5�@�ǉ��@�i���`�I�����@����~�����@�΍�j
'                                      'E'�̏����Ɂ@genten�@�ǉ�
'            update: 2004. 3. 8 s.f.  �ύX�@���`�����䃂�[�h�@�f�V�f�ǉ��@�i�㎲�Փ˔���t�j
'                                    RecEmgDTsave ����~���b�Z�[�W�̕ۑ�
'            update: 2004. 3.12 s.f.  ���x�w�ߓd���@�\��
'
'            update: 2004. 4.23 s.f.  timeup�Ŕ���~
'            update: 2004. 4.24 s.f.  �J�E���^�A�����сA�\���@����
'            update: 2004. 5. 5 s.f   ���x�W���A�����␳���[�`���@�ǉ�  PGM_KTD,My_lib,MYEDIT, LS21_SC, LS21_TC
'            update: 2004.5.12  s.f   PGM_KTD�@"���ް�۰"�΍�@�@wTm0!,wTm1!  global��,  LS21_SC�Ɓ@LS21_TC ����@dim�폜
'            update: 2004.5.17  s.f   'S'����ށ@�o�O�΍�
'            update: 2004.5.18  s.f   �o�O�΍� & T�W���\��
'            update: 2004.8.17  s.f   ���ް�۰"�΍�  p(ist0)��pp��  �h�F�h�����̍s�𖳂���
'            update: 2004.8.27 - 10.30 s.f   T�W���֐��ύX�A�@�@�u�c�b�@�O�v�R�}���h�@���`�O�Ɍ^�ݔۃ`�F�b�N�Z���T�[�̃`�F�b�N�@�\�ǉ�
'            update: 2004.12.20 s.f   D�b�@�O�v�R�}���h�@�o�O�C��
'            update: 2005. 5.25 s.f    Version No�\���ǉ�
'            update: 2005. 7.18 s.f    �ŏI���`�I����@�Q�O���̎��R��p
'            update: 2005. 9.28 s.f   T�W���@�\���F�ύX
'            update: 2005.11.22 s.f   Melec C-870 counter����o�O�C���@�R���y�A�J�E���^�l�Z�b�g���@�������]�@�@setcm1
'                                     C870sts(3) ����@�o�O�C��, �E���f�[�^�\�������ύX
'            update: 2005.11.23 s.f   11/22 �ύX�̃o�O�C���@���`������@�uC870sts�@reset����܂Ł@�ǂݔ�΂��v���@����
'            update: 2005.12.17 s.f   Do-Loop �O�́@DoEvent�폜 OverFlow �΍� s.f.
'                                     �R�}���h�́@evtime�@��荞�݂��@�R�}���h�J�n���֕ύX
'�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@DC�R�}���h�@LA�R�}���h�@�ă`�F�b�N�C��
'�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�A���O�R�}���h�@evtime�@�Ɓ@fintime�@�\�L����ւ�
'            update: 2006. 3. 3 s.f  edit �g�p���@do�@loop���甲����
'�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@DC����ނց@fintime=timer�@���@�ݒu
'            update: 2006. 4.14 s.f  on error goto ,  sts as long
'            update: 2006. 4.15 s.f  error �\��
'            update: 2006. 5. 9 s.f  O.F.error �\���@������@end3�@�ǉ�,  tstime=0#
'            update: 2006. 5.18 s.f �@r_pres()�́@DoEvents �@�폜�A�@�hJ"�A�P�b��1��@Doevents�@�ǉ�
'                                    ����~�@�\���ǉ�
'       Ver.3.33R_061221 2006.12.21 s.f  LS-33���@�Ή��@�@VacuumON�AVacuumOFF�@��p�~�ASeikeiON,SeikeiOFF�V�݁@DO3�@���蓖�ĕύX
'       Ver.3.33R_070927 2007.09.27 s.f  Z�␳�@�w�肵��������No.�ց@�ł���悤�ɂ���
'           '
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Dim lGphNo%
Dim lGphNo0%
Dim EditFlg%
Dim lViewFlg      '�O�̉�ʔԍ�
Dim NextView%
Dim lEmgFlg%      '����~
Dim lDtSaveFlg%   '�f�[�^�ۑ�
Dim TCFlg%        '�e�X�g���`��
'Dim iPltMax%      '�p���b�g��]��    '05.7.18 global��
Dim l_stime!      '�T�C�N���^�C��
Dim lHO_Flg%      'HO�R�}���h�p�t���O
Dim lHO_Time!     'HO�R�}���h�̎���
'Dim lSokuFlg%     '�����\�[�N�^�C��
Dim CmndCol!(0 To 1)  '�R�}���h�t�̐F
Dim SokuCor!(0 To 1)  '�����\�[�N�^�C���̃R�}���h�t�̐F
'Dim T_keisuCol!(0 To 1)  '���x�W���A�����␳�\����backColor
Dim lCycleTime$       '�T�C�N���^�C��
'Dim sdt1$, sdt2$, sdt3$   2006.4.14 global he
'Dim iFlg_hijyou%    '����~�t���O  s.f. 2004.3.8   2009.8.17�폜
Dim TCi_s%         ' �u�P�񐬌`�v���́@���`��
Private Sub Command2_Click(Index As Integer)
Select Case Index
Case 0  '�L�����Z��
  lGphNo = 0
  MoniGraph Me.Picture1, 0, lGphNo
Case 1  '�I��
  If TCFlg = True Then         '�e�X�g���`��
    FrmMenuFlg = False
    NextView = 1
  Else
    Unload Me
    PGM_Menu.Show
  End If
Case 2  '�O���t�ĕ`��
  lGphNo = lGphNo + 100
  MoniGraph Me.Picture1, 0, lGphNo
Case 3  'edit �G�f�B�^�N��
  Unload Me
  MYEdit.Show
Case 4      '�^�󓞒B
  gVumFlg = 1                       '�^�󓞒B=1
Case 5      '"S" ;�f�[�^�Z�[�u
  lDtSaveFlg = True
Case 6      '���`�J�n
  iPltMax = 3    '�p���b�g��]��
  Timer1.Enabled = False
  Command2(1).Caption = "���f"
  Command2(3).Enabled = False
  Command2(6).Enabled = False
  Command2(7).Enabled = False
  TC_Main
  Command2(3).Enabled = True
  Command2(6).Enabled = True
  Command2(7).Enabled = True
  Command2(1).Caption = "�I��"
  Timer1.Enabled = True
Case 7      '���`�J�n
  iPltMax = Val(Text1(0))    '�p���b�g��]��
  Timer1.Enabled = False
  Command2(1).Caption = "���f"
  Command2(3).Enabled = False
  Command2(6).Enabled = False
  Command2(7).Enabled = False
  TC_Main
  Command2(3).Enabled = True
  Command2(6).Enabled = True
  Command2(7).Enabled = True
  Command2(1).Caption = "�I��"
  Timer1.Enabled = True
Case 8      '�����\�[�N�^�C��
  If lSokuFlg = True Then
          lSokuFlg = False          '�����\�[�N�^�C���@��t����
          Command2(8).BackColor = SokuCor(0)
    Else
          lSokuFlg = True           '�����\�[�N�^�C���@��t
          Command2(8).BackColor = SokuCor(1)
  End If
End Select

End Sub

Private Sub SetData()

Dim l_sdt$

  Label2(0) = Format(ptime, "###0")  '���莞��
  Label2(1) = Format(ytemp, "###0")  '�\�����M���x
  Label2(2) = gcoxFlName             '����t�@�C����
  Label2(3) = hcomm(2)               '�R�����g
  '
'  Label2(13).Caption = Str(InitDat(11))   '���`�J�E���^�g�E�^��      TC_main ���ŏ���
'/* ���Đ�������јg�\�� */
'  l_sdt = Format(l_stime / 60, "0") & "��" & Format(Int(l_stime) Mod 60, "0") & "�b"    '2002.10.09 KYOCERA
'  Label2(9).Caption = Format(InitDat(10), "000")    '���`�J�E���^ i_s
'  Label2(8).Caption = l_sdt               '�^�N�g�^�C��
' -----------------------------------
  DispGphScale
End Sub

Private Sub GetData()

End Sub

Private Sub Form_Load()
  DispCenter Me
  LS21_TC.Caption = LS21_TC.Caption + "     " + versionNo
  Me.Top = 0
  SokuCor(0) = &H8000000F     '�����\�[�N�^�C���̃R�}���h�t�̐F
  SokuCor(1) = &HFF&          '�����\�[�N�^�C���̃R�}���h�t�̐F �����ꂽ�Ƃ�
'  T_keisuCol!(0) = &HFFFFFF    '���x�W���A�����␳�@�\��backcolor�@off
'  T_keisuCol!(1) = &HFFFFC0    '���x�W���A�����␳�@�\��backcolor�@on
  lDtSaveFlg = False      '�f�[�^�ۑ�
'  'lSokuFlg = False        '�����\�[�N�^�C��
  If lSokuFlg = False Then
          Command2(8).BackColor = SokuCor(0)
    Else
          Command2(8).BackColor = SokuCor(1)
  End If
  lViewFlg = ViewFlg      '�O�̉�ʔԍ�
  ViewFlg = 3             '��ʔԍ�
  FrmMenuFlg = True       '���j���[���甲����Ƃ�false
  EditFlg% = False        '�G�f�B�^�N������
  lEmgFlg = False         '����~
  TCFlg = False           '�e�X�g���`��
  Command2(1).Caption = "�I��"
  SetData
  TrnsReqON               '�����˗��M���n�m
  Timer1.Enabled = True
End Sub
Private Sub Form_Unload(Cancel As Integer)
  HeatOFF       '/* �\�����M���[���ɂ��A�n�e�e���� */
  CoolOFF
  ServoOFF
  TrnsReqOFF    '�����˗��M��OFF
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
  'Label2(4).Caption = ""
  'Timer1.Enabled = False
  '-------------- �s���j�v�ǂ�
'  LS21S_Monitor        '2006.12.21 �폜 s.f
  'LS21T_MAIN
End Sub
Public Sub LS21T_MAIN()   '/* �P�񐬌` ���C���v���O���� 2002.5.28a*/
'/*�@�@�@�@�@�@�@�@�@�@�@�@�@�@���x�\���ǉ��A�@���ԕ\���@�ǉ��@2002.6.15�@*/
Dim i%, j%, js%, l%, ist0%, ist1%, ndata!, mdata!, ntemp!, mtemp!, ntemp0!, mtemp0!, iflg%, isflg%
Dim ied%, ips%, i_s%, irei%, r_ch%, ix%, ix0%, iy%, isp%, stime%, ii%, iii%, istend%
Dim ie02%, ie03%, ie04%, ituflg%, iSRcount
Dim ie%, ie0%, ie1%, ie2%, ie3%, ie4%, ie5%
Dim isub As Long, jsub As Long, ksub As Long, lsub As Long
Dim sdata!, sv%, zch%     '  05.11.26 s.s. overflow �΍�
'Dim sdata%, sv%, zch%
Dim ct_dummy!, iz3%, itc%
'Dim m_l%, sdata%, sv%, zch%
Dim com$, tdate$, ttime$
Dim m_l!
Dim st!, ev!, sev!, fin!, it!          '/* ���ԗp�f�[�^ */
Dim btemp!(0 To 4), bposi!, bpre! '/* ���x�@�ʒu�@���� �̑O�f�[�^ */
'
Dim sdt$, ch%, hdt%, flindex%, imax%, sts1%, sts2%, ch1%, ch2%
Dim sts As Long        ' 2006.4.14
Dim S_StartTime!
Dim stTime!, evtime!, sevTime!, mTime!, tsTime!, endTime!
'Dim stTime!, evtime!, fintime!, sevTime!, mTime!, tsTime!, endTime!
Dim dt!(0 To 7)
Dim flg As Long, cnt As Long
Dim diTime!, diTime1!, diTime2!, diTimeSR!, pdt!, pp!, pml!, x1dt!, x2dt!, pos!
Dim dmy$, idmy%, iwt!
Dim r_z_now!, r_z_ave!, r_z_dum!(0 To 180), it_ts%, i_ts%    ' /* 2002.4.9�@�ǉ��@�˓����`�@*/
Dim epsilon!
Dim cp_z!, cc_time!(0 To 3), ct_temp!(0 To 2), ct!, ict%   ' CP , CT �p
Dim ct_t!(0 To 10)
'
 On Error GoTo errHandler:
  iFlg_hijyou = 0
  iz3 = 3   '  Z(ist0)�@�́@�@Z3�́@index�l
  ips = 1
  i_s = -1              '���`��
  For ii = 1 To 180: r_z_dum(ii) = 0#: Next ii
  tsTime = 0#
'  �ݒ�ʒu���ֈړ�
'  it_ts = roz(1)       ' 10     '/* �˂����ĒB���@�����@���ς���� mzx 180 */
'  epsilon = roz(0)     ' 0.0005 '/* �˓��@���e���@�@mm�@�@*/
'
'----------------------- �P�񐬌`���C���v���O����
  ppos = "TC"     ' ���݈ʒu= TC
  C870Stop
  ServoON       '/* �T�[�{���� */
  CtlDisp       '�ʒu����
  'TrnsReqOFF    '�����˗��M��OFF   SC�̎�
  TrnsReqON      '�����˗��M���n�m�@TC�̎�
'/***********     �گ��@C-853�{�[�h�����ݒ�@�@�@*************/
'/* SPEC INITIALIZE CMD OUT */
'/* �J�E���^�{�[�h�̏����ݒ� */
  InitDat(10) = 0
'/* ������ڰľ�ĺ���� */
  C870AccRate
'/* ���x�ݒ� */
'  C870LSPDSet 300    '/* 300 pps 0.066mm/sec */
  C870LSPDSet 800    '/* 300 pps 0.066mm/sec */
'/* �f�B���[�^�C���ݒ� */
  C870DelayTime
  rstcm1   '  compareter reset
'/***********     �گ��@C-853�{�[�h�����ݒ�@�I��  *************/

'/* �`�s�b���x���Z�b�g */
'/* ���{�b�g�f�[�^�̃t���b�s�[����̓ǂ݂Ƃ� */
  rozFileLoad
'
  it_ts = Int(roz(1))  ' 10       '/* �˂����ĒB���@�����@���ς���� max180*/
  epsilon = roz(0)    ' 0.0005   '/* �˓��@���e���@�@mm�@�@*/
'
st:
  If ied = 2 Then GoTo st2:
'/*  ����t�@�C���̃I�[�v�� */
  coxDtRead gcoxFldir & gcoxFlName
  Label2(0).Caption = Format(ptime, "0")
  If T_keisuCont(2) <> 0 Then T_keisuCont(1) = T_keisuCont(2) '�|�C���^�[backup
  If T_keisuCont(3) <> 0 Then T_keisuCont(0) = T_keisuCont(3) '�^�� backup
'/* �O���t�B�b�N��ʂ̏����� */
  InitDat(8) = ptime  '�O���t�X�P�[���o�ߎ��� (Max)
  SetData
  lGphNo0 = 0
  lGphNo = 0
  MoniGraph Me.Picture1, lGphNo0, lGphNo
'/* ���x�W���@�����␳�̕\�� */
  For itc = 0 To 9
    Label4(itc).Caption = Format(T_keisu(itc), "0.000")
    Label6(itc).Caption = Format(Z3_Hosei(itc), "0.000")
    If itc < T_keisuCont(0) Then
       Label4(itc).BackColor = T_keisuCol!(1)
       Label6(itc).BackColor = T_keisuCol!(1)
      Else
        Label4(itc).BackColor = T_keisuCol!(0)
        Label6(itc).BackColor = T_keisuCol!(0)
    End If
  Next itc
'
'/* ���쓮����R�}���h�̃t�@�C������̓ǂݎ�� */
  i = 0
  Do
    Label2(12).Caption = DispCtrlCommand(i)
    If pres(i) >= 1000 Then ips = 2         '/* ��ڽ����1ton�ȏ�Ŏ��ύX */
    i = i + 1                               '/* ips�͎������`�掞�̃X�P�[��para*/
    If ic(i - 1) = 9 Then Exit Do
  Loop

  ic(i) = 10                                ' /* �ŏI�Z�O�����g�{�P�Ɂ@�u10�v���Z�b�g */
  istend = i                           ' ���������ނ�end�ԍ�
'ic(i) = 4
'/* �\��̕\�� */
  Label2(2).Caption = gcoxFlName
'/* ���_�o�� */
  Label2(4).Caption = "���_�o�����s"
  genten
  Ready_Wait
  Counter0
  Label2(4).Caption = "���_�o������"
'/* �J�E���^�Ƀ[������������ */
  'C870AdrInit       '�`�c�c�q�d�r�r �h�m�h�s�`�k�h�y�d �b�n�l�l�`�m�c
  C870CntPreSet 0   '�b�n�t�m�s�d�q �o�q�d�r�d�s �b�n�l�l�`�m�c
  'InitDat(10) = 0
  pos = r_z()
  GCnt0 = 0
  GCnt1 = 0
'/* �����^�]�F�� */
  Label2(4).Caption = "�����^�]�F����"
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
  Label2(4).Caption = ""
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
    If com = "S" Then
      jsub = sjsub(flindex)
      ksub = sksub(flindex)
      lsub = slsub(flindex)
      sdt = sdt & Right(Space(15) & Format(jsub, "0"), 15)
      sdt = sdt & Right(Space(15) & Format(ksub, "0"), 15)
      sdt = sdt & Right(Space(15) & Format(lsub, "0"), 15)
    Else
      
    End If
    Label2(7).Caption = sdt
    flindex = flindex + 1
    i = 10
    '
    If ied <> 0 Then GoTo jp0:
    '
    Select Case com
      Case "B"
        Label2(4).Caption = "CASE B DO1"
        'Exit Do
      Case "N"    '/* ���f�K�X�̒��� */
        If Mid(scom(flindex), 2, 1) = "S" Then
          If isub = 1 Then
            Label2(4).Caption = "���f�K�X���� DO1"
            N2Open
          End If
          If isub = 0 Then
            Label2(4).Caption = "���f�K�X��~ DO1"
            N2Close
          End If
        End If
      Case "J"    '/* ���ԑ҂� */
        Label2(4).Caption = "���ԑ҂� DO1"
        evtime = Timer
        Do
          fintime = Timer2func     ' 2009.8.17
'          fintime = Timer
          DoEvents
          Label2(10).Caption = Format(diffTime(fintime, evtime), "0")
          If diffTime(fintime, evtime) >= isub Then Exit Do
        Loop
        Label2(10).Caption = ""
      Case "K"    '/* ���M */
        Select Case Int(isub)
        Case 1
          Label2(4).Caption = "���M�@�n�m DO1"
          HeatON
        Case 0
          Label2(4).Caption = "���M�@�n�e�e DO1"
          HeatOFF
        End Select
      Case "S"    '/* �`�s�b���x�ݒ� */
        Label2(4).Caption = "�`�s�b���x�ݒ� DO1"
        evtime = Timer              '�҂����߂̎���
        ntemp0 = isub
        mtemp0 = jsub
        ntemp0 = T_keisu_cset(ntemp0, T_keisu(T_keisuCont(1) - 1))  'ntemp0
        mtemp0 = T_keisu_cset(mtemp0, T_keisu(T_keisuCont(1) - 1))  'mtemp0
        Do
          fintime = Timer2func     ' 2009.8.17
'          fintime = Timer
          DoEvents
          diTime = diffTime(fintime, evtime)
          If ksub <> 0 Then x1dt = diTime / ksub
          ndata = (ntemp0 - ntemp) * x1dt + ntemp
          mdata = (mtemp0 - mtemp) * x1dt + mtemp
          TempSet 2, ndata
          TempSet 3, mdata
          If diTime >= ksub Then Exit Do
        Loop
        ntemp = ntemp0
        mtemp = mtemp0
        TempSet 2, ntemp
        TempSet 3, mtemp
      Case "R"    '/* ��p */
        Select Case Int(isub)
        Case 0    '��p��@�n�e�e
          Label2(4).Caption = "��p��@�n�e�e DO1"
          CoolOFF
        Case 1    '��p��@�n�m
          Label2(4).Caption = "��p��@�n�m DO1"
          CoolON
        Case 2    '��p���@�n�m
          Label2(4).Caption = "��p���@�n�m DO1"
          CoolON
        End Select
    End Select
jp0:
    If i < 24 Then
      i = i + 1
    Else
      Label2(4).Caption = ""
    End If
    If com = "B" Then Exit Do
  Loop
'/* ���`�v���Z�X�A���^�]�J�n */
'/* �f�[�^��ǂݎ�� */

'/* �u�U�[��炷 */
  'Label2(4).Caption = ""
  'Label2(10).Caption = ""
st2:
'/* �^�C�g���̕\�� */
'/* �^�������̕\�� */
'/* ���W�l���̕\�� */
'/* �����p�y���ʒu�ύX�g�\�� */
'  Label2(5).Caption = Format(roz(0), "0.0000")     '/* �˓����`para�@���@ */�@02.10.26 s.f �폜
'  Label2(6).Caption = Format(roz(1), "0.0")     '/* �˓����`para�@���� */�@02.10.26 s.f �폜
'/* ���`�J�n */
  Do        '----------------- DO LOOP
    DoEvents
    i_s = i_s + 1                   ' /* i_s = ���`�� */
    js = 0
    ist0 = -1
    ist1 = -1           '/* ist0 ist1�@(�����l -1) ���������ނ̌��ݔԍ� */
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
    Label4(T_keisuCont(1) - 1).BorderStyle = 1    '  �g�t���ɂ���(�|�C���^�[�j
    Label6(T_keisuCont(1) - 1).BorderStyle = 1  '  �g�t���ɂ���(�|�C���^�[�j
    iz3 = Z3_HoseiCont(2)   ' Z�␳�@�����{����@ZNo.�@�@�@�f07.9.27�@�ǉ�
    z(iz3) = z(iz3) + Z3_Hosei(T_keisuCont(1) - 1) '  �hZ3"�̕␳�lset
'/* �J�E���^�ւ̏o�͂��� */
'                                             TC_main �Ŏ��{
'    If i_s <> 0 Then
'      InitDat(11) = InitDat(11) + 1  '���`�J�E���^�g�E�^��
'      InitDtSave
'      Label2(13).Caption = Str(InitDat(11))   '���`�J�E���^�g�E�^��
'    End If
'/* ���`�g�̕\�� */
ejs1:
  lGphNo0 = 0
  lGphNo = 0
  MoniGraph Me.Picture1, lGphNo0, lGphNo
'/* �w���̕\�� */
'/* �x���̕\�� */
'/* ���Đ�������јg�\�� */
    sdt = Format(stime / 60, "0") & "��" & Format(Int(stime) Mod 60, "0") & "�b"        '2002.10.09 KYOCERA
'    Label2(9).Caption = Format(i_s, "000")
'    Label2(8).Caption = sdt         '�T�C�N���^�C��
    lCycleTime = sdt                '�T�C�N���^�C��
    InitDat(10) = i_s               '���`�J�E���^
'
'    For iii = 0 To 9
'       Label6(iii).Caption = ""
'    Next iii
'
'/* �J�E���^�ւ̏o�̓_�E�� */
    'InitDat(11) = InitDat(11) - 1   '���`�J�E���^�g�E�^��
    'InitDtSave
    'Label2(13).Caption = Str(InitDat(11))
'/* �f�[�^�̎�荞�� */
'    stTime = Timer            DO loop �J�n����ց@�ړ��@10/5
    evtime = Timer
    iflg = 1
    ied = 0
    ttime = Time
    mTime = Timer
'----------------------------- For Loop i
    imax = ptime * 60
    For i = 1 To imax
start:
    'finTime = Timer
'    DoEvents           '2005.12.17  s.f.
    ituflg = 0            '�@�^�C���A�b�vflg�̃��Z�b�g10/5
'/* ���`���̃h���C�u*/
      If ist0 > 0 Then
       'If ic(ist0 - 1) = 4 Then     '/* ������I��������а������ */
        If ic(ist0 - 1) = 10 Then    '/* �u�ŏI������+1�v�́A�@�u10�v*/
          ist0 = ist0 - 1            '/* ���[�v�f�ʂ�̂��߂���а */
        End If
      End If
        sdt3$ = DispSegm(ist0)
        Label2(12).Caption = sdt3$
      If ist0 <> ist1 Then
        gOrgFlg = False                '���_���A����=TRUE
        ist1 = ist0
        sevTime = Timer '            �J�n���Ԃ̎�荞��
'
        If (ist0 > 0 And ist0 < 11) Then   '�@�J�n���Ԃ̕\���@�����������p
           diTime1 = diffTime(sevTime, stTime)          '2002.10.09 KYOCERA
           sdt = Format(ist0, "0") & "=" & Format(Int(diTime1 / 60), "0") & "'" & Format(Int(diTime1) Mod 60, "0") & "�b"       '2002.10.09 KYOCERA
'           Label6(ist0 - 1).Caption = sdt
        End If
'
        Select Case ic(ist0)  '-------- �����䃂�[�h�ԍ�
        Case 0, 8   '-------------------- �ʒu����
          ppos = "TC JikuStart 0 8"
          Ready_Wait    '
          CtlDisp     'outp(DIO_P+3,9); �T�[�{ON & ���x���S12
          s_drive z(ist0), vel(ist0)
        Case 1, 3, 7   '-------------------- ���x����    2004.3.8 �u7�v�ǉ�
          ppos = "TC JikuStart 1 3 7"
          m_l = vel(ist0)
          'm_l = vel(ist0) / 100
          If m_l > 50 Then m_l = 50
          setcm1 z(ist0)
          Ready_Wait    '
          CtlVelo       'outp(DIO_P+3,5);
          'while((inp(XCN_DT1)&0x01)!=0);
          Do    '' �u�J�E���^�[��v�v��ԒE�o�p
            DoEvents
            sts = C870Sts(3)    'sts=1�̎��@���������u-1�v�@sts=0�̎��s���������u0�v
            If (sts And &H1) = 0 Then Exit Do   'PULSE �� COMPARE ����v
          Loop
          '
        Case 2    '-------------------- �_�~�[
          ppos = "TC JikuStart 2"
          Ready_Wait
          CtlDisp  'DioOut 12,1  �ʒu���� '  02.10.1 �ǉ�
          Ready_Wait    '
          ServoON     'outp(DIO_P+3,1);
        Case 9    '-------------------- �I��
          ppos = "TC JikuStart 9"
          Ready_Wait    '
          CtlDisp     'outp(DIO_P+3,9);
          genten
          'Ready_Wait
          For ii = 1 To 180         '/* ����R�p�̂̏����� */
            r_z_dum(ii) = 0#
          Next ii
          i_ts = 1
          r_z_ave = 0#
        End Select
      End If
'
        fintime = Timer2func     ' 2009.8.17
'        fintime = Timer         '2002.10.09 KYOCERA

'/* �^�C���A�b�v���� */
          '2002.10.09 KYOCERA
      If ist0 < 0 Then GoTo sj1:
      'If ituflg = 0 Then
          If ((ic(ist0) < 10) And (diffTime(fintime, sevTime) > t0(ist0))) Then  '2002.10.16 KYOCERA 2002.10.17 KYOCERA            '10/4
            ituflg = 1
            sdt = "�^�C���A�b�v" & Right(Space(11) & Format(diffTime(fintime, sevTime), "0.0"), 11)
            sdt = sdt & Right(Space(11) & Format(t0(ist0), "0.0") & Format(ist0 + 1, "0"), 11)
            Label2(5).Caption = sdt + "TUp=" + Str(gTimeUpCnt) & Str(ist0) & "  ����;" & Format(Now, "hh:mm:ss")
'
                RecEmgDtSave sdt3, sdt1, sdt2
                gemgmsg = "��ѱ���"
                hijyou        '����~����
                iFlg_hijyou = 1        '�@��ѱ���
                GoTo eend:
'
            ist0 = ist0 + 1             '/�^�C���A�b�v�Ŏ��̃X�e�b�v   '2002.10.16 KYOCERA
            'GoTo TimeUpEnd:
            GoTo jscmdend:              '�@�I���M���������щz��    10/12 sf
          End If
      'Else                          ' �_�u���`�F�b�N�@�P�D�Q�b��ɍĊm�F
          'If ((ic(ist0) < 9) And (diffTime(finTime, sevTime) > (t0(ist0) + 1.2))) Then            '10/4
            'sdt = "�^�C���A�b�v" & Right(Space(11) & Format(diffTime(finTime, sevTime), "0.0"), 11)
            'sdt = sdt & Right(Space(11) & Format(t0(ist0), "0.0") & Format(ist0 + 1, "0"), 11)
            'gTimeUpCnt = gTimeUpCnt + 1    '�^�C���A�b�v�̃J�E���^
            'label2(5).Caption = sdt + "TUp=" + Str(gTimeUpCnt) & Str(ist0)
            'ist0 = ist0 + 1             '/�^�C���A�b�v�Ŏ��̃X�e�b�v
            'hijyou        '����~����
            'getch
            'GoTo eend:
            'ituflg = 0
            'GoTo jscmdend:              '�@�I���M���������щz��    10/4 sf
          'End If
      'End If
TimeUpEnd:
'
'/* �I���M���̏��� */
      Select Case ic(ist0)
      Case 0, 8   '/* �ʒu����̏ꍇ */
          ppos = "TC JkE0 8"
        If (C870Sts(1) And 1) = 0 Then
           Label2(5).Caption = "�ʊ�sg=" & Str(ist0 + 1)  '���No.=ist0+1 10/4  sf
           ist0 = ist0 + 1
        End If
      Case 1    '/* ���x����̏ꍇ */
          ppos = "TC JkE1"
        pdt = pres(ist0)
        pp = p(ist0)
        pml = m_l
          ppos = "TC JkE1 -1cal"
        cal_pid pdt, pp, pml
          ppos = "TC JkE1 cal_pid"
        sts = C870Sts(3)  'status3 ��ǂ�
        If (sts And &H1) <> 0 Then
          ist0 = ist0 + 1             '/* �ʒu�B���ŏI�� */
            Label2(5).Caption = "�ʒu���� pass CNT " & Str(ist0)    '11/2 s.f
            rstcm1   '  compareter reset
            Ready_Wait    '
         Else                       ' 2008.2.21  �ύX�@�P�b�ɂP��s���߂����m�F��
'
           If Int(mTime) = Int(Timer) Then
             If r_z() >= z(ist0) Then
               ist0 = ist0 + 1             '
               Label2(5).Caption = "�ʒu pass PC " & Str(ist0)
             End If
           End If
         End If
         ppos = "TC JkE1 r_z -1"
'        Else
'          If r_z() >= z(ist0) Then
'            ist0 = ist0 + 1             '
'            Label2(5).Caption = "�ʒu���� pass PC " & Str(ist0)    '11/2 s.f
'          End If
'          ppos = "TC JkE1 r_z -1"
'        End If
      Case 3    '/* ���x����@�˓����`�̏ꍇ  2002.4.9 */
          ppos = "TC JkE3"
        pdt = pres(ist0)
        pp = p(ist0)
        pml = m_l
          ppos = "TC JkE3 -1cal"
        cal_pid pdt, pp, pml
          ppos = "TC JkE3 cal_pid"
        sts = C870Sts(3)  'status3 ��ǂ�
          ppos = "TC JkE3 sts=C870"
        If (sts And &H1) <> 0 Then
          ist0 = ist0 + 1             '/* �ʒu�B���ŏI�� */
          Label2(5).Caption = "�ʒu���� pass CNT " & Str(ist0)    '11/2 s.f
            rstcm1   '  compareter reset
            Ready_Wait    '
         Else                       ' 2008.2.21  �ύX�@�P�b�ɂP��s���߂����m�F��
'
           If Int(mTime) = Int(Timer) Then
             If r_z() >= z(ist0) Then
               ist0 = ist0 + 1             '
               Label2(5).Caption = "�ʒu pass PC " & Str(ist0)
             End If
           End If
         End If
         ppos = "TC JkE3 r_z -1"
'
'        Else
'          If r_z() >= z(ist0) Then
'            ist0 = ist0 + 1             '
'            Label2(5).Caption = "�ʒu���� pass PC " & Str(ist0)    '11/2 s.f
'          End If
'            ppos = "TC JkE3 r_z -1"
'        End If
'        If r_z() < z(ist0) Then
'          r_z_now = r_z()
'          If Int(tsTime) <> Int(mTime) Then
'              ppos = "TC JkE3 r_z -2"
'            tsTime = mTime                  '/* �P�b�O�ƁA�Q�b�O�� */
'            If Abs(r_z_now - r_z_ave) < epsilon Then
'              ist0 = ist0 + 1               '/* it_ts��A���@epsilon�ȉ� */
'            Else                            '/* �Ł@�˓��B���ŏI�� */
'              r_z_dum(i_ts) = r_z_now
'              r_z_ave = 0#
'              For ii = 1 To it_ts
'                 r_z_ave = r_z_ave + r_z_dum(ii)
'              Next ii
'              r_z_ave = r_z_ave / it_ts
'              i_ts = i_ts + 1
'              If i_ts > it_ts Then i_ts = 1
'            End If
'          End If
'        End If
'
'
          If r_z() < z(ist0) Then
'            r_z_now = r_z()                    '2008.2.23 �ړ�
              ppos = "TC JkE3 r_z -2"
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
'�@�@�@�@�@�@�@/*�@�w�舳�͂�荂�����͂��R�b�ȏ㑱���������~�@�@*/
        ppos = "TC JkE7"
        pdt = pres(ist0)
        pp = p(ist0)
        pml = m_l
        cal_pid pdt, pp, pml
        sts = C870Sts(3)  'status3 ��ǂ�
        If (sts And &H1) <> 0 Then
          ist0 = ist0 + 1             '/* �ʒu�B���ŏI�� */
          Label2(5).Caption = "�ʒu pass CNT " & Str(ist0)   '11/2 sf
            rstcm1   '  compareter reset
            Ready_Wait    '
'          Do            '����do�@loop�@�Ӗ������@�@2005.11.22  s.f
'            DoEvents
''            sts = C870Sts(3)          'status3 ��ǂ�    10/4  sf
''            If (sts And &H1) <> 0 Then Exit Do           10/4 sf
'             If r_z() >= z(ist0) Then Exit Do             '10/4
'          Loop
         Else                       ' 2008.2.21  �ύX�@�P�b�ɂP��s���߂����m�F��
'
           If Int(mTime) = Int(Timer) Then
             If r_z() >= z(ist0) Then
               ist0 = ist0 + 1             '
               Label2(5).Caption = "�ʒu pass PC " & Str(ist0)
             End If
           End If
         End If
'
'        Else
'          If r_z() >= z(ist0) Then
'            ist0 = ist0 + 1             '
'            Label2(5).Caption = "�ʒu pass PC " & Str(ist0)
'          End If
'        End If
'        bpre = r_pres()
'        If bpre > pdt Then              ' 2008.2.18  �ύX
'        If bpre > pdt * 0.7 Then
'          If Int(tsTime) <> Int(mTime) Then
'            tsTime = mTime                  '/* �P�b�O�Ɣ�r */
'            i_ts = i_ts + 1               '/* i_ts��A�����ā@���͂��w��l�ȏ� */
'
'
        If Int(tsTime) <> Int(mTime) Then '2008.2.23 �ύX 1�b��1��`�F�b�N
          tsTime = mTime                  '/* �P�b�O�Ɣ�r */
          bpre = r_pres()
          If bpre > pdt Then                ' 2008.2.18 �ύX
'               If bpre > pdt * 0.7 Then
             i_ts = i_ts + 1               '/* i_ts��A�����ā@���͂��w��l�ȏ� */
             If i_ts > 3 Then
                gemgmsg = "������7 error"
                hijyou        '����~����
                'getch
                iFlg_hijyou = 2     '������@7�@error
                GoTo eend:
             End If
          End If
        End If                                 '/*     '2004.3.8�@�����܂Ł@*/
      Case 9    '�I��
        ppos = "TC JkE9"
        sts = C870Sts(1)
        If (sts And 1) = 0 Then
          ist0 = ist0 + 1     '/* ���� */
          If Abs(r_z()) > 0.1 Then
            Label2(5).Caption = "���_�s��"
            ist0 = ist0 - 1
            genten              '���_�o��
          End If
        Else
          '/* �J�E���^�Ƀ[������������ */
          Ready_Wait
          Counter0
        End If
      End Select
'
'      Select Case ic(ist0)           ' 2004.3.12 s.f  05.11.26 nuku
'        Case 1, 3, 7
'                Label8(0).Caption = nout
'                Label8(1).Caption = v
'      End Select
'
jscmdend:                               '������R�}���h�@������  10/4 sf
'
'/* �G���[�\�� */
    If ArmChk <> 0 Then               '�A���[�����b�Z�[�W
      frmerr_sign.Show   'ALM�o��
    Else
      Unload frmerr_sign
    End If
    
'    If ArmChk <> 0 Then   '�A���[�����b�Z�[�W�@�@'03.7.10��L�ɕύX
'      frmerr_sign.Show 1�@�@�@�@�@�@�@�@�@�@�@�@'03.7.10��L�ɕύX
'    End If�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@�@'03.7.10��L�ɕύX

'/* �v���Z�X���s */
sj1:
    If iflg = 1 Then
      com = scom(js + flindex)
      isub = sisub(js + flindex)
      jsub = sjsub(js + flindex)
      ksub = sksub(js + flindex)
      lsub = slsub(js + flindex)
      js = js + 1
'
      evtime = Timer        ' 2005.12.17  s.f.  �J�n���Ԏ�荞��
'
      sdt = com & Right(Space(7) & Format(isub, "0"), 7)
      'locate(60,1);printf("%2s %5ld ",com,isub);
      If Left(com, 1) = "S" Then
        sdt = sdt & Right(Space(7) & Format(jsub, "0"), 7)
        sdt = sdt & Right(Space(7) & Format(ksub, "0"), 7)
        sdt = sdt & Right(Space(7) & Format(lsub, "0"), 7)
        'printf("%5ld %5ld",jsub,ksub);
        'Label2(7).Caption = sdt
      Else
        sdt = sdt
      End If
      Label2(7).Caption = sdt
    End If
        '�V�X�e�����f�B/* ����~�̏ꍇ�͐��`���~ */
        sts1 = SystemReadyChk()   '�V�X�e�����f�B or ����~
        sts2 = AutoChk()          '������ԁH (<>0 ����)
        If sts1 = 0 Or sts2 = 0 Then
          Label2(4).Caption = "�����^�]�I��"
          gemgmsg = ArmEmgMsgChk$()
          iFlg_hijyou = 10              '����~ү���ނ̂�������
          FrmMenuFlg = False              '���j���[���甲����Ƃ�false
          NextView = 1
          Exit Do
        End If
        '
        Select Case Left(com, 1)
'
        Case "D"    '------------ ���`���̌^�̗L��
            ppos = "TC Proc D"
            If (isub = 0) Then     '�ݔۃZ���T�[�@�`�F�b�N
              If (KataChk() > 0) Then                '  2004.10.30  �^�ݔۃ`�F�b�N�p�Z���T�̓���m�F�p
                 sdt = "DC�@�ݔۃZ���T�[�ُ�i�^�L��I�I�j"
                 Label2(5).Caption = sdt
'
                  RecEmgDtSave sdt3, sdt1, sdt2
                  gemgmsg = "DC error �^�L��"
                  hijyou        '����~����
                  iFlg_hijyou = 3     '   DC error  �^�L��
                  GoTo eend:
              Else
                GoTo scend:
              End If
            End If                           '  2004.10.30  �^�ݔۃ`�F�b�N�p�Z���T�̓���m�F�p
'
           If KataChk() < 3 Then '�^������
            Label2(4).Caption = "CASE D ���`���^���� DO2"
            fintime = Timer2func     ' 2009.8.17
'            fintime = Timer    '���ݎ��ԁ@�@2006.3.3�@�ǉ��@s.f.
            If (diffTime(fintime, evtime) < isub) Then
               iflg = 0             ' ���Ԗ��B�̏ꍇ
            Else
               idmy = js            '�@���ԑ҂��I���̏ꍇ
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
                         gemgmsg = "DC error 4"
                         hijyou        '����~����
                         iFlg_hijyou = 4     '"DC error 4"
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
                          gemgmsg = "DC error 5"
                        hijyou        '����~����
                         iFlg_hijyou = 5     '"DC error 5"
                         GoTo eend:
'
                   End If
                 Loop
'
              iflg = 1                    '�@����ޏI������
              i_s = i_s - 1             '���`�񐔂̖߂�
'              evtime = Timer          '2005.12.17  s.f.
            End If
          End If                    '�@�^������ꍇ�͂��̂܂ܔ�����
'
        Case "L"    ' ���`���Ɍ^�������������̔�ѐ�@��а
            ppos = "TC Proc L"
            Label2(4).Caption = "CASE L ���`���^���� DO2"
          'iflg = 1�@�@�@�@�@�@�펞�@iflg=1�@�ׁ̈A�K�v�Ȃ�
'
        Case "H"    ' �����\�[�N�@�@�@�h�g�b�h
            ppos = "TC Proc H"
            Label2(4).Caption = "�����\�[�N DO2"
            fintime = Timer2func     ' 2009.8.17
'            fintime = Timer    '���ݎ��ԁ@�@2006.3.3�@�ǉ��@s.f.
            If (lSokuFlg = True And diffTime(fintime, evtime) < isub) Then
               iflg = 0
             Else
               iflg = 1
               lSokuFlg = False
               Command2(8).BackColor = SokuCor(0)
'               evtime = Timer          '2005.12.17  s.f.
            End If
'
        Case "S"    '/* �`�s�b���x�ݒ� */
            ppos = "TC Proc S"
            Label2(4).Caption = "�`�s�b���x�ݒ� DO2"
'            DoEvents          '2005.12.17  s.f.�폜�@�@2006.5.18�ǉ��@�폜
            If Mid(com, 2, 1) = "R" Then             ' SR�̏ꍇ  ���F�֘A�������@Do�@Loop�@Top�ɂ���
              fintime = Timer2func     ' 2009.8.17
'              fintime = Timer
               diTime = diffTime(fintime, stTime)    ' 0.1�b�ɂP�񉷓x��荞�݁i�T����{�j
               If ((diTime - diTimeSR) > 0.1) Then
'                   ct_t(0) = ct_t(0) + TempRdMold() '���x�Ǎ�
                   ct_dummy = TempRdMold(0)    '�X���[�u���x�Ǎ�
                   ct_dummy = T_keisu_cread(ct_dummy, T_keisu(T_keisuCont(1) - 1))
                   ct_t(0) = ct_t(0) + ct_dummy '���x�Ǎ�
                   iSRcount = iSRcount + 1
                   diTimeSR = diTime
                   iflg = 0
                   If iSRcount > 5 Then
                      ct_t(0) = ct_t(0) / 5
                      ntemp0 = isub
                      ntemp0 = T_keisu_cset(ntemp0, T_keisu(T_keisuCont(1) - 1)) 'ntemp0
                      mtemp0 = jsub
                      mtemp0 = T_keisu_cset(mtemp0, T_keisu(T_keisuCont(1) - 1)) 'mtemp0
                      ntemp0 = ct_t(0) + ntemp0
                      mtemp0 = ct_t(0) + mtemp0
                      ntemp = ntemp0
                      mtemp = mtemp0
                      TempSet 2, ntemp
                      TempSet 3, mtemp
                      ct_t(0) = 0
                      Label2(5).Caption = "SR= " & Format(Int(ntemp), "000") & Format(Int(mtemp), "  000")
                      iSRcount = 1
                      iflg = 1
'                      evtime = Timer          '2005.12.17  s.f.
                   End If
               End If
            Else
              fintime = Timer2func     ' 2009.8.17
'              fintime = Timer
              diTime = diffTime(fintime, evtime)        'SA�̏ꍇ
              If ksub <> 0 Then x1dt = diTime / ksub
              ntemp0 = isub
              mtemp0 = jsub
              ntemp0 = T_keisu_cset(ntemp0, T_keisu(T_keisuCont(1) - 1))  'ntemp0
              mtemp0 = T_keisu_cset(mtemp0, T_keisu(T_keisuCont(1) - 1))  'mtemp0
              ndata = (ntemp0 - ntemp) * x1dt + ntemp
              mdata = (mtemp0 - mtemp) * x1dt + mtemp
              TempSet 2, ndata
              TempSet 3, mdata
              If diTime >= ksub Then
                iflg = 1
                ntemp = ntemp0
                mtemp = mtemp0
                TempSet 2, ntemp
                TempSet 3, mtemp
'                evtime = Timer          '2005.12.17  s.f.
              Else
                iflg = 0
              End If
            End If
'
        Case "P"    '/* �ړ�������̋쓮 */
            ppos = "TC Proc P"
            Label2(4).Caption = "�ړ�������̋쓮 DO2"
          If Mid(com, 2, 1) = "W" Then
            Beep
            ist0 = ist0 + 1
            sevTime = Timer          '2005.12.17  s.f.
'           evtime = Timer          '2002.10.09 KYOCERA          '2005.12.17  s.f.
          End If
          If Mid(com, 2, 1) = "R" Then
            iflg = 0
            If ist0 <> ist1 Then iflg = 1
            If isub = 4 And ist0 = 0 Then iflg = 1
'            If iflg = 1 Then evtime = Timer         '2002.10.09 KYOCERA          '2005.12.17  s.f.
          End If
          'evTime = Timer
        Case "K"    '/* ���M */
          ppos = "TC Proc K"
          Select Case isub
          Case 1
            Label2(4).Caption = "���M ON DO2"
            HeatON
          Case 0
            HeatOFF
            Label2(4).Caption = "���M OFF DO2"
          End Select
        Case "N"
            ppos = "TC Proc N"
            Label2(4).Caption = "CASE N DO2"
          If Mid(com, 2, 1) = "S" Then
            If isub = 1 Then hdt = hdt
            If isub = 0 Then hdt = hdt
          End If
        Case "R"    '/* ��p */
          ppos = "TC Proc R"
          Select Case isub
          Case 2
            Label2(4).Caption = "��p ON1 DO2"
            CoolON
          Case 1
            Label2(4).Caption = "��p ON2 DO2"
            CoolON
          Case 0
            Label2(4).Caption = "��p OFF DO2"
            CoolOFF
          End Select
        Case "T"    '/* �`�s�b�P�̉��x�̓ǂݎ�� */
            ppos = "TC Proc T"
            Label2(4).Caption = "�`�s�b�P�̉��x�̓ǂݎ�� DO2"
          sdata = TempRdMold(0)    '�X���[�u���x
          sdata = T_keisu_cread(sdata, T_keisu(T_keisuCont(1) - 1)) 'ndata
          If (Mid(com, 2, 1) = "L" And sdata > isub) Or (Mid(com, 2, 1) = "G" And sdata < isub) Or (Mid(com, 2, 1) = "E" And (sdata > (isub + 20) Or sdata < (isub - 20))) Then
'          If (Mid(com, 2, 1) = "L" And sdata > isub) Or (Mid(com, 2, 1) = "G" And sdata < isub) Then
            iflg = 0
          Else
            If iflg = 2 Then iflg = 1 Else iflg = 2
'            evtime = Timer          '2005.12.17  s.f.
          End If
        Case "J"    '/* ���ԑ҂� */
          ppos = "TC Proc J"
          DoEvents      '2006.5.18  s.f �ǉ�
            Label2(4).Caption = "���ԑ҂� DO2"
            fintime = Timer2func     ' 2009.8.17
'            fintime = Timer    '���ݎ��ԁ@�@2006.3.3�@�ǉ��@s.f.
          diTime1 = diffTime(fintime, stTime)
          diTime2 = diffTime(fintime, evtime)
          If (Mid(com, 2, 1) = "S" And diTime1 >= isub) Or (Mid(com, 2, 1) = "C" And diTime2 >= isub) Then
            iflg = 1
'            evtime = Timer          '2005.12.17  s.f.
          Else
            iflg = 0
          End If
'
        Case "C"
          ppos = "TC Proc C"
          Select Case Mid(com, 2, 1)
          Case "P"    '���`�I���ʒu�@�`�F�b�N
            Label2(4).Caption = "���`�I���ʒu�@�`�F�b�N DO2"
            cp_z = r_z()
            Label5(0).Caption = " cp=" & Format(cp_z, "0.000")
'
          Case "C"    '�@���ԃ`�F�b�N
            Label2(4).Caption = "���ԃ`�F�b�N DO2"
            If isub > 3 Then
                ict = 5
              Else
                ict = isub + 2
            End If
            fintime = Timer2func     ' 2009.8.17
'            fintime = Timer    '���ݎ��ԁ@�@2006.3.3�@�ǉ��@s.f.
            cc_time(isub) = diffTime(fintime, stTime)
            sdt = " cc" & Format(isub, "0") & "= " & Format(Int(cc_time(isub) / 60), "0") & "��" & Format(Int(cc_time(isub)) Mod 60, "0") & "�b"        '2002.10.09 KYOCERA
            Label5(ict).Caption = sdt
            If isub = 3 Then
                diTime1 = diffTime(cc_time(isub), cc_time(isub - 1))
                sdt = " cc3-2=  " & Format(Int(diTime1 + 0.5), "0") & "�b"
                Label5(6).Caption = sdt
            End If
'
          Case "T"    '�@���x�`�F�b�N
            Label2(4).Caption = "���x�`�F�b�N DO2"
            If isub > 2 Then
                ict = 2
              Else
                ict = isub
            End If
            ct_temp(isub - 1) = TempRdMold(0) '���x 0V-300�� 1V-1300��
            ct_temp(isub - 1) = T_keisu_cread(ct_temp(isub - 1), T_keisu(T_keisuCont(1) - 1))
            sdt = " ct" & Format(isub, "0") & "=" & Format(ct_temp(isub - 1), "0.0") & "��"
            Label5(ict).Caption = sdt
          End Select
'
        Case "X"    '�����I���M���i���`�J�n�j
          ppos = "TC Proc X"
          Select Case Mid(com, 2, 1)
          Case "R"    '���`�J�n [�����I���܂ő҂�]
            Label2(4).Caption = "���`�J�n [�����I���܂ő҂�] DO2"
            '--------------------- TC�ō폜
            'TrnsReqON  '�����˗��M��Ch21�o��
            '
            'Do
              '-------------- �s���j�v�ǂ�
            '  LS21S_Monitor
              'DioInput 13, sts        '�����I���H
            '  sts = TrnsFinChk()      '�����I���H
            '  If sts = 1 Then
            '    TrnsReqOFF            '�����˗��M���n�e�e
            '    Label2(4).Caption = "�����˗��M���n�e�e"
            '    Exit Do
            '  End If
            '  DoEvents
            'Loop
            '--------------------- TC�ō폜
          Case "W"    '���`�I��
            Label2(4).Caption = "���`�I�� DO2"
          End Select
        Case "E"    '/* �I���@���{�b�g���� */
            ppos = "TC Proc E"
           If r_z() > 2 Then                                      '03.9.11
              genten                                              '03.9.11
              'Ready_Wait    'while((inp(AX_STS)&1)!=0);          '03.9.11
            End If
             Label2(4).Caption = "�I�� ���{�b�g���� DO2"
        '--------------------- TC�ō폜
           iflg = 99
           GoTo send:
        '  Exit Do
scend:
        End Select
cjump:
  '-------------- �s���j�v�ǂ�
'  LS21S_Monitor      '2006.12.21 �폜 s.f
  lEmgFlg = SystemReadyChk()  '����~
  If Int(mTime) = Int(Timer) And lEmgFlg <> 0 Then GoTo start:
  'lEmgFlg = EmgChk()         '����~
  'If Int(mTime) = Int(Timer) And lEmgFlg = False Then GoTo start:
'                 /* 1�b��1�񉺂ɔ����� */
      mTime = Timer
      ppos = "TC 1 sec Disp 1"
'
    If FrmMenuFlg = False Then             '���j���[���甲����Ƃ�false
      Select Case NextView
      Case 1
        sdt = "�I������t"
      Case 8  'edit
        sdt = "edit����t"
      Case Else
      
      End Select
      Label2(10).Caption = sdt
    Else
      Label2(10).Caption = ""
    End If
'           /* ���́@�o�h�c����@�o���P�T�@�Ȃ瑬�x�@�[�� */
  If ist0 >= 0 Then
    If p(ist0) > 15 Then
      DaVoltOut 1, 0        ' 0V D/A ch=1
    End If
  End If
'/*�@���x��荞�� */
'    DoEvents          '2005.12.17  s.f.
    atemp(i, 0) = TempRdMold(0)   '�X���[�u���x 0V-300�� 1V-1300��
    atemp(i, 0) = T_keisu_cread(atemp(i, 0), T_keisu(T_keisuCont(1) - 1))
    atemp(i, 1) = 0                 '�ヂ�[���h���x
    atemp(i, 1) = T_keisu_cread(atemp(i, 1), T_keisu(T_keisuCont(1) - 1))
    atemp(i, 2) = 0                 '�����[���h���x
    atemp(i, 2) = T_keisu_cread(atemp(i, 2), T_keisu(T_keisuCont(1) - 1))
  
'* ���`���ʒu�̎�荞�� */
      ppos = "TC 1 sec Disp 2"
      aposi(i) = r_z()
      '
'/* �^���͂̎�荞�� */
      ppos = "TC 1 sec Disp 3"
      apre(i) = r_pres()

'      If i = 1 Then GoTo jo0:
'      ix0 = Int(8.3333 / ptime * (i - 1)) + 120
'      ix = Int(8.3333 / ptime * (i)) + 120
'-------------- �s���j�v�ǂ�
'      LS21S_Monitor
'/* ���x���z�̕\�� */
'/* �^�����̃v���b�g */
'/* ���W�l�̃v���b�g */
    lGphNo = i
    GphDataSet lGphNo0, lGphNo
    MoniGraph Me.Picture1, lGphNo0, lGphNo
    lGphNo0 = lGphNo
jo0:
'/* �e��f�[�^�̉�ʉ��\�� �P�@*/
    DoEvents     '  2006.5.18  �ǉ�
    sdt1 = Right(Space(10) & Format(atemp(i, 0), "0.0"), 10) & "��"
    sdt1 = sdt1 & Right(Space(10) & Format(apre(i), "0.00"), 10) & "kgf"
    sdt1 = sdt1 & Right(Space(10) & Format(aposi(i), "0.000"), 10) & "mm"
    Label2(14).Caption = sdt1
'/* �e��f�[�^�̉�ʉ��\�� �Q */
    it = Timer                                                          ' 10/5
    it = diffTime(it, stTime)
    sdt2 = Right(Space(2) & Format(Int(it / 60), "0"), 2) & "��"
    sdt2 = sdt2 + Right(Space(2) & Format(Int(it) Mod 60, "0"), 3) & "�b�@"       '2002.10.09 KYOCERA
    sdt2 = sdt2 + "ct" & Right(Space(7) & Format(diffTime(fintime, evtime), "0.0"), 7) & "  "
    sdt2 = sdt2 + "st" & Right(Space(7) & Format(diffTime(fintime, sevTime), "0.0"), 7) & "  "
    sdt2 = sdt2 + "tt" & Right(Space(7) & Format(diffTime(fintime, stTime), "0.0"), 7)
    Label2(11).Caption = sdt2
'/* �����\�� */
    Label10.Caption = Time$
'
'/* ��ޯĈʒu�ύX�@*/
    'If FrmMenuFlg = False Then GoTo eend:
  Next i   '--------------------------------- For Loop
  js = js - 1
  GoTo ejs1:      '/* �\���I���Ō���ʂ� */
'/* �^�N�g�^�C���̎Z�o�@*/

send:
      ppos = "TC 1��end"
 
 '   stime = i
'    endTime = Timer
'    stime = diffTime(endTime, stTime)         '  10/5
'    sdt = Format(Int(stime / 60), "0") & "��" & Format(Int(stime) Mod 60, "0") & "�b"   '2002.10.09 KYOCERA
'    lCycleTime = sdt
'    Label2(6).Caption = Format(stime, "000") & Format(i_s, " 000")         '02.10.26 s.f. �폜
'/* �f�[�^�̕ۑ��@*/
    If lDtSaveFlg = True Then
      ResDtSave i_s, stime
      lDtSaveFlg = False
    End If
'�@/*�@���`�f�[�^�̃Z�[�u�@*/  2002.12.4 sf
'      Rec_of_Mold = Format(InitDat(11), "000") & "  "
      Rec_of_Mold = "   " & Format(z(iz3), "000.00") & "  "
      Rec_of_Mold = Rec_of_Mold & "  " & Format(Int(ct_temp(0)), "000") & "�� " & Format(Int(ct_temp(1)), "000") & "�� "
      Rec_of_Mold = Rec_of_Mold & "     " & Format(Int(cc_time(1) / 60), "0") & ":" & Format(Int(cc_time(1)) Mod 60, "00") & " "
      Rec_of_Mold = Rec_of_Mold & "  " & Format(Int(cc_time(2) / 60), "0") & ":" & Format(Int(cc_time(2)) Mod 60, "00") & " "
      Rec_of_Mold = Rec_of_Mold & "  " & Format(Int(cc_time(3) / 60), "0") & ":" & Format(Int(cc_time(3)) Mod 60, "00") & " "
      diTime1 = diffTime(cc_time(3), cc_time(2))
      Rec_of_Mold = Rec_of_Mold & "  " & Format(Int(diTime1 + 0.5), "000") & "s "
      Rec_of_Mold = Rec_of_Mold & "    " & Format(cp_z, "000.000")
      Rec_of_Mold = Rec_of_Mold & "    " & Format(Int(stime / 60), "0") & ":" & Format(Int(stime) Mod 60, "00") & " "
      Rec_of_Mold = Rec_of_Mold & "    " & Format(T_keisu(T_keisuCont(1) - 1), "0.000") & "    " & Format(Z3_Hosei(T_keisuCont(1) - 1), "0.000") & "  "

'    Rec_of_Mold = Format(InitDat(11), "000") & "  "�@�@' TC_MAIN �Ŏ��{
'
'    Rec_of_Mold = " z " & Format(z(iz3), "000.00") & "  " & Format(z(5), "000.00") & " "
'    Rec_of_Mold = Rec_of_Mold & " :  ct " & Format(Int(ct_temp(0)), "000") & "�� " & Format(Int(ct_temp(1)), "000") & "�� "
'    Rec_of_Mold = Rec_of_Mold & " :  cc " & Format(Int(cc_time(1) / 60), "0") & ":" & Format(Int(cc_time(1)) Mod 60, "00") & " "
'    Rec_of_Mold = Rec_of_Mold & "  " & Format(Int(cc_time(2) / 60), "0") & ":" & Format(Int(cc_time(2)) Mod 60, "00") & " "
'    Rec_of_Mold = Rec_of_Mold & "  " & Format(Int(cc_time(3) / 60), "0") & ":" & Format(Int(cc_time(3)) Mod 60, "00") & " "
'    diTime1 = diffTime(cc_time(3), cc_time(2))
'    Rec_of_Mold = Rec_of_Mold & " :  " & Format(Int(diTime1 + 0.5), "0") & "s "
'    Rec_of_Mold = Rec_of_Mold & " : cp   " & Format(cp_z, "000.000")
'    Rec_of_Mold = Rec_of_Mold & " : t    " & Format(Int(stime / 60), "0") & ":" & Format(Int(stime) Mod 60, "00") & " "
'
'    RecDtSave Rec_of_Mold    ' TC_MAIN �Ŏ��{
' /* ���x�W���A�����␳�f�[�^�̃J�E���g�A�b�v
      Label4(T_keisuCont(1) - 1).ForeColor = T_keisuCol!(2)  '  �����F�����ɖ߂�
      Label6(T_keisuCont(1) - 1).ForeColor = T_keisuCol!(2)  '  �����F�����ɖ߂�
      Label4(T_keisuCont(1) - 1).BorderStyle = 0  '  �g�Ȃ��ɖ߂�
      Label6(T_keisuCont(1) - 1).BorderStyle = 0  '  �g�Ȃ��ɖ߂�
'     *** Z3�̒l���@�߂�
          z(iz3) = z(iz3) - Z3_Hosei(T_keisuCont(1) - 1) '  �hZ3"�̕␳�lreset
'     *** �|�C���^�[�J�E���g�A�b�v
      T_keisuCont(1) = T_keisuCont(1) + 1       ' �|�C���^�[�̃J�E���g�A�b�v ������
'      Z3_HoseiCont(1) = Z3_HoseiCont(1) + 1       ' �|�C���^�[�̃J�E���g�A�b�v
    If T_keisuCont(1) > (T_keisuCont(0)) Then T_keisuCont(1) = 1
'
    T_keisuCont(2) = T_keisuCont(1)       ' �|�C���^�[��buckup
    T_keisuCont(3) = T_keisuCont(0)       '  �^�� backup
'
'    If Z3_HoseiCont(1) > (Z3_HoseiCont(0)) Then Z3_HoseiCont(1) = 1
'
'/* �u���͂�����Ă�����@�G�f�B�b�g */
    If FrmMenuFlg = False Then Exit Do            '���j���[���甲����Ƃ�false
    If EditFlg% = True Then '�G�f�B�^�N��
       ied = 1
       Exit Do
    End If
'/* ������~��Ԃł���Β�~ */
'    sts1 = SystemReadyChk()   '�V�X�e�����f�B or ����~
'    sts2 = AutoChk()          '������ԁH
'    If sts1 = 0 Or sts2 = 0 Then
      Label2(4).Caption = "�����^�]�I��"
'      FrmMenuFlg = False              '���j���[���甲����Ƃ�false
'      NextView = 1
      Exit Do
'    End If
  Loop    '------------------------------------ DO LOOP
'/* �u���͂�����Ă�����@�G�f�B�b�g */
    If ied = 1 Then '�G�f�B�^�N��
       MYEdit.Show 1
       'c = 0
       ied = 0
       GoTo st:     '/* �G�f�B�b�g���[�h�ł���΁@�����ɃW�����v */
    End If
'/* �G�f�B�b�g���[�h�ł���΁@�����ɃW�����v */
'    If ied <> 0 Then GoTo st:
'/* �\�����M���[���ɂ��A�n�e�e���� */
eend:
    If iFlg_hijyou > 0 Then
         RecEmgDtSave sdt3, sdt1, sdt2 & gemgmsg
   End If                 '����~���b�Z�[�W�̕ۑ�  2004.3.8
  HeatOFF
  CoolOFF
'  ServoOFF
'/* ���{�b�g�f�[�^�̃t���b�s�[�ւ̏����o�� */
'/* �O���t�B�b�N��ʂ̏��� */

'/* �u���͂�����Ă�����@�G�f�B�b�g */
Exit Sub
'
errHandler:
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

'Private Sub cal_pid(m_sa!, m_p!, m_lim!)
'  float  m_sa,     /* �ݒ舳�� */
'         m_p,      /* �ݒ�o�l */
'         m_lim;    /* �ݒ胊�~�b�g�l */
'Dim i%, nout%, ch%, v!
'Dim pa!, per!       '/* float�i�P���x���������_�^)*/
'  pa = r_pres()     '/* ���� */
'
''  If pa > 1000# Then '/* 1000�j���ȏ�Ŕ���~ */
'  If pa > m_sa + 200# Then '/* �w�舳�́{200�j���ȏ�Ŕ���~ */
'  hijyou
'    Exit Sub
'  End If
'
''/* �o�h�c���Z */
'
'  per = 5 * (m_sa - pa) * Abs(m_sa - pa) / (m_p * m_p)
'  If per > m_lim Then per = m_lim
'  If per < (-1 * m_lim) Then per = -1 * m_lim
'  'nout = Int(40.95 * per) + &H800
'  nout = &H800 - Int(4.095 * per / 4#)
'  'nout = &H800 - Int(40.95 * per)
'  ch = 1
'  v = per / 5
'  'v = per / 5
'  DaOut ch, Hex(nout)
'  'DaVoltOut ch, V
'  'outp(ADPORT,(nout%256));
'  'outp(ADPORT+1,0x20|(nout/256));
'
'End Sub

Private Sub GphXSet()
Dim i%
  For i = 0 To ptime * 60 + 10
    TPass(i) = i
  Next i
End Sub

Private Sub GphDataSet(i0%, i1%)
Dim i%
  For i = i0 To i1
    Templ(i) = atemp(i, 0)    '�����v
    Templu(i) = atemp(i, 1)   '��^
    Templd(i) = atemp(i, 2)   '���^
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


Private Sub TC_Main()
Dim i%, flg%
Dim TCstTime!, TCendTime!, TCstime!
'------------ ���`�I�����[
  SeikeiON         '���`ON�@�A�����͂P�񐬌`��
  TCFlg = True          '�e�X�g���`��
  lEmgFlg = SystemReadyChk()  '����~
  lCycleTime = "------"         '�T�C�N���^�C��
  InitDat(11) = InitDat(11) - 1  '���`�J�E���^�g�E�^�����킹
  TCstTime = Timer
  TCi_s = -1
'
'/* ���`�f�[�^�ۑ��t�@�C���̍쐬�@*/
  RecDtSave0 InitDat(11)
'
'----------
  For i = 1 To iPltMax     '�p���b�g��]��
    If lEmgFlg = 0 Or FrmMenuFlg = False Then Exit For
    Label2(4).Caption = "�p���b�g" & Trim(Str(i)) & "����"
    TCi_s = TCi_s + 1
    Label2(9).Caption = TCi_s   '���`�J�E���^
'
   Plt1Jyun
'
    If lEmgFlg = 0 Or FrmMenuFlg = False Then Exit For
    If i <> iPltMax Then
'
'  ------ ���`�J�E���^�Ǘ� -------
          InitDat(11) = InitDat(11) + 1  '���`�J�E���^�g�E�^��
          InitDtSave
          Label2(13).Caption = Str(InitDat(11))   '���`�J�E���^�g�E�^��
' /* ---  ���`���C�� ---
      LS21T_MAIN
'
'    ---  �����с@�@�v�Z
        TCendTime = Timer
        TCstime = diffTime(TCendTime, TCstTime)
        lCycleTime$ = Format(Int(TCstime / 60), "0") & "��" & Format(Int(TCstime) Mod 60, "0") & "�b"
        Label2(8).Caption = lCycleTime$           '�T�C�N���^�C��
        TCstTime = Timer
'
      Rec_of_Mold = Format(i, "000") & "  " & Rec_of_Mold  '���`�f�[�^��save
      RecDtSave Rec_of_Mold
'
      If iFlg_hijyou = 1 Or lEmgFlg = 0 Or FrmMenuFlg = False Then Exit For
'
    End If
  Next i
  TCFlg = False         '�e�X�g���`�I��
  SeikeiOFF          '���`OFF�@�ҋ@��
  If lEmgFlg <> 0 Then
    If FrmMenuFlg = False Then
      Label2(4).Caption = "���f"
      FrmMenuFlg = True
    Else
      coolingform.Show
'
      WaitSec (1)
      flg = MsgBox("���`�E��p�@�I�� " + Time$ + "   ", 48, "1�񐬌`") '�I�����b�Z�[�W
    End If
  Else
    RecEmgDtSave sdt3, sdt1, sdt2         '����~���b�Z�[�W�̕ۑ�  2004.3.8
'
    Unload Me
    PGM_Menu.Show
  End If
End Sub

Private Sub Plt1Jyun()
Dim sts%
'---------- �p���b�g1���w�߁�1�������܂ő҂�
  WaitSec 1.5
  PCTrnsReq     ' �p���b�g1���w��
  
  '2002.10.9 KYOCERA
  sts = 0
  Do
    sts = PCTrnsChk()   'BUSY�`�F�b�N
    lEmgFlg = SystemReadyChk()  '����~
    If sts = 1 Or lEmgFlg = 0 Then Exit Do
    DoEvents
  Loop
  
  sts = 0
  Do
    sts = PCTrnsChk()   'PC���������=1
    lEmgFlg = SystemReadyChk()  '����~
    If sts = 0 Or lEmgFlg = 0 Then Exit Do
    DoEvents
  Loop
End Sub

Private Sub Timer2_Timer()
If r_z > 0.1 Then
        OrgOFF
    Else
        OrgON
    End If
End Sub
