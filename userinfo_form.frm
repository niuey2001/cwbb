VERSION 5.00
Begin VB.Form userinfo_form 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "用户信息维护"
   ClientHeight    =   2445
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   2445
   ScaleWidth      =   4560
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   2355
      Left            =   0
      ScaleHeight     =   2295
      ScaleWidth      =   4425
      TabIndex        =   0
      Top             =   0
      Width           =   4485
      Begin VB.CommandButton Command2 
         Cancel          =   -1  'True
         Caption         =   "退出"
         Height          =   435
         Left            =   2040
         TabIndex        =   4
         Top             =   1560
         Width           =   945
      End
      Begin VB.CommandButton b_ok 
         Caption         =   "确定"
         Default         =   -1  'True
         Height          =   435
         Left            =   600
         TabIndex        =   3
         Top             =   1560
         Width           =   945
      End
      Begin VB.TextBox text_psw 
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
         Top             =   840
         Width           =   1635
      End
      Begin VB.TextBox text_user_name 
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
         Top             =   240
         Width           =   1635
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
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
         TabIndex        =   7
         Top             =   1440
         Width           =   120
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "密码"
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
         TabIndex        =   6
         Top             =   840
         Width           =   450
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "用户名"
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
         TabIndex        =   5
         Top             =   240
         Width           =   675
      End
   End
End
Attribute VB_Name = "userinfo_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()


End Sub

Private Sub b_ok_Click()
Dim rs As ADODB.Recordset
Dim user_name As String
Dim user_psw As String
user_psw = Me.text_psw
Dim strsql
Dim t As VbMsgBoxResult
    Set rs = New ADODB.Recordset
    Set rs.ActiveConnection = cn
    rs.LockType = adLockOptimistic
    rs.CursorType = adOpenKeyset
    strsql = "select * from t_user_info where user_name='" & userName & "' and password = '" & text_psw & "'"
    rs.Open strsql
    If Not rs.EOF Then
        If Trim(user_name) <> "" Then
        
            rs.Fields("user_name") = user_name '是否有问题 ？？
            rs.Update
        Else
            MsgBox "用户名不可为空！"
        End If
        
    End If
    



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
End Sub

Private Sub Command2_Click()
Unload Me

End Sub

