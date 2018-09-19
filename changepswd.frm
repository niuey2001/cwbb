VERSION 5.00
Begin VB.Form change_psw_form 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "修改密码"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5535
   Icon            =   "changepswd.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   5535
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   2985
      Left            =   120
      ScaleHeight     =   2925
      ScaleWidth      =   5025
      TabIndex        =   0
      Top             =   120
      Width           =   5085
      Begin VB.TextBox txtold 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   960
         Width           =   1635
      End
      Begin VB.TextBox txtnew 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   1560
         Width           =   1635
      End
      Begin VB.TextBox txtsec 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   300
         IMEMode         =   3  'DISABLE
         Left            =   1560
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   2160
         Width           =   1635
      End
      Begin VB.CommandButton b_ok 
         Caption         =   "确定"
         Default         =   -1  'True
         Height          =   435
         Left            =   3480
         TabIndex        =   4
         Top             =   1650
         Width           =   945
      End
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "退出"
         Height          =   405
         Left            =   3480
         TabIndex        =   5
         Top             =   2190
         Width           =   945
      End
      Begin VB.Label lab_name 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   1560
         TabIndex        =   10
         Top             =   360
         Width           =   1695
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用 户 名"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   915
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "旧 密 码"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   8
         Top             =   960
         Width           =   915
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "新 密 码"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   150
         TabIndex        =   7
         Top             =   1560
         Width           =   915
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "确定密码"
         BeginProperty Font 
            Name            =   "楷体_GB2312"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   120
         TabIndex        =   6
         Top             =   2160
         Width           =   900
      End
   End
End
Attribute VB_Name = "change_psw_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_ok_Click()
Dim rs As ADODB.Recordset
Dim strsql As String

Dim oldPassword As String

Dim t As VbMsgBoxResult

Call check_condatabase
Set rs = New ADODB.Recordset
Set rs.ActiveConnection = cn
strsql = "select * from t_user_info where user_name='" & username & "'"
rs.Open strsql, cn, adOpenKeyset, adLockOptimistic
If Not rs.EOF Then
    oldPassword = rs("password")
    'MsgBox oldPassword
   If txtold.Text = oldPassword Then
     If txtnew.Text = txtsec.Text Then
        If txtnew.Text = "" Or txtsec.Text = "" Then
           t = MsgBox("请输入密码!", 48, "WARNING")
        Else                             '如满足条件，则更新密码
        rs.Fields("password") = txtnew.Text '是否有问题 ？？
        rs.Update
        t = MsgBox("密码修改成功！", vbOKOnly, "SURE")
        txtold.Text = ""
        txtnew.Text = ""
        txtsec.Text = ""
        End If
      Else
      t = MsgBox("密码输入不同！", 48, "warning")
      txtnew.Text = ""
      txtsec.Text = ""
      End If
  Else
    t = MsgBox("原密码错误！", 48, "warning")
    txtold.Text = ""
    txtnew.Text = ""
    txtsec.Text = ""
  End If

End If

End Sub

Private Sub Command2_Click()
Unload Me

End Sub


Private Sub name_Click()

End Sub

