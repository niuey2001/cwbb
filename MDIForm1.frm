VERSION 5.00
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   7185
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10635
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000009&
      Height          =   21535
      Left            =   0
      Picture         =   "MDIForm1.frx":0000
      ScaleHeight     =   21480
      ScaleWidth      =   19020
      TabIndex        =   0
      Top             =   0
      Width           =   19080
      Begin VB.TextBox text_name 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         TabIndex        =   4
         Top             =   3360
         Width           =   1695
      End
      Begin VB.TextBox text_password 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   4680
         TabIndex        =   3
         Top             =   4200
         Width           =   1695
      End
      Begin VB.CommandButton Command1 
         BackColor       =   &H00FF8080&
         Caption         =   "新用户注册"
         BeginProperty Font 
            Name            =   "华文新魏"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   3840
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5400
         Width           =   2055
      End
      Begin VB.CommandButton button_login 
         BackColor       =   &H00FF8080&
         Caption         =   "登陆"
         BeginProperty Font 
            Name            =   "华文新魏"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6360
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   5400
         Width           =   1575
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "山西重点税源监控财务指标采集客户端"
         BeginProperty Font 
            Name            =   "华文新魏"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   975
         Left            =   720
         TabIndex        =   7
         Top             =   480
         Width           =   8055
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "用户名："
         BeginProperty Font 
            Name            =   "华文新魏"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3000
         TabIndex        =   6
         Top             =   3360
         Width           =   1935
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "密   码："
         BeginProperty Font 
            Name            =   "华文新魏"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3000
         TabIndex        =   5
         Top             =   4200
         Width           =   1575
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Picture1_Click()

End Sub
