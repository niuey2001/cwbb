VERSION 5.00
Begin VB.Form login_form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "财务指标采集客户端------山西省地税局"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9675
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6585
   ScaleWidth      =   9675
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H80000009&
      Height          =   21535
      Left            =   0
      Picture         =   "Form1.frx":16AC2
      ScaleHeight     =   21480
      ScaleWidth      =   9615
      TabIndex        =   0
      Top             =   0
      Width           =   9675
      Begin VB.ComboBox Combo1 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   4080
         TabIndex        =   8
         Top             =   2880
         Width           =   1695
      End
      Begin VB.CommandButton b_login 
         BackColor       =   &H00FF8080&
         Caption         =   "登陆"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   4
         Top             =   4440
         Width           =   2055
      End
      Begin VB.CommandButton b_register 
         BackColor       =   &H00FF8080&
         Caption         =   "新用户注册"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   6840
         MaskColor       =   &H00FF8080&
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   5280
         Width           =   2055
      End
      Begin VB.TextBox text_password 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   4080
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   3750
         Width           =   1695
      End
      Begin VB.TextBox text_name 
         BorderStyle     =   0  'None
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   6600
         TabIndex        =   1
         Top             =   1080
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "密   码："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   615
         Left            =   2400
         TabIndex        =   7
         Top             =   3720
         Width           =   1575
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "用户名："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H8000000D&
         Height          =   615
         Left            =   2400
         TabIndex        =   6
         Top             =   2880
         Width           =   1935
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "山西重点税源监控财务指标采集客户端"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   975
         Left            =   960
         TabIndex        =   5
         Top             =   360
         Width           =   8055
      End
   End
End
Attribute VB_Name = "login_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub b_login_Click()
Dim userNameRs As ADODB.Recordset
Dim pdnsrd As ADODB.Recordset

Dim v_name As String  '用户名
Dim v_password As String  '密码

v_name = Combo1.Text
v_password = text_password.Text

Call check_condatabase
sql = "select user_name,user_type from t_user_info where user_name = '" & v_name & "' and password = '" & v_password & "'"
Set userNameRs = cn.Execute(sql)
If Not userNameRs.EOF Then
     'MsgBox "登陆成功！"
    username = v_name   'username为全局变量  即登陆成功后的用户名  任意窗口都可取到
    userType = userNameRs("user_type")
    Unload Me
    userNameRs.Close
    sql = "select * from t_nsrxx where username='" & v_name & "'"
    Set pdnsrd = cn.Execute(sql)
     If Not pdnsrd.EOF Then
     MainForm.Show
     main_form.BorderStyle = none
     main_form.Show
     Else
     imponsrxx_form.Show
    
     End If
    
    
    
    'operate_form.BorderStyle = none
    'operate_form.Show
   
Else
   MsgBox "用户名或密码有误！"
End If

Set userNameRs = Nothing



'If v_name = array_userInfo(0) And v_password = array_userInfo(1) Then
'    'MsgBox "登陆成功！"
'    Unload Me
'    MainForm.Show
'    operate_form.BorderStyle = none
'    operate_form.Show
'
'Else
'    MsgBox "用户名或密码有误！"
'End If


End Sub

Private Sub b_register_Click()
register_form.text_password = ""
register_form.text_password_two = ""
register_form.text_username = ""
register_form.Show
login_form.Hide
register_form.text_username.SetFocus
End Sub

Private Sub Form_Load()
Dim sql As String
Dim userRs As ADODB.Recordset
Me.text_name = ""
Me.text_password = ""
 Call condatabase  '连接数据库
sql = "select * from t_user_info "
Set userRs = cn.Execute(sql)
 While Not userRs.EOF
        If Trim(userRs("user_name")) <> "" Then
        Me.Combo1.AddItem userRs("user_name")
        End If
        userRs.MoveNext
    Wend
    
    If Combo1.ListCount > 0 Then
    Me.Combo1.ListIndex = 0
    End If

End Sub

