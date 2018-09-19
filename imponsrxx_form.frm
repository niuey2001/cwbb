VERSION 5.00
Begin VB.Form imponsrxx_form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "纳税人信息录入"
   ClientHeight    =   6555
   ClientLeft      =   45
   ClientTop       =   540
   ClientWidth     =   9645
   ForeColor       =   &H00FF0000&
   Icon            =   "imponsrxx_form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "imponsrxx_form.frx":16AC2
   ScaleHeight     =   6555
   ScaleWidth      =   9645
   Begin VB.ComboBox combo_hy 
      Height          =   300
      ItemData        =   "imponsrxx_form.frx":2C05E
      Left            =   3840
      List            =   "imponsrxx_form.frx":2C060
      Style           =   2  'Dropdown List
      TabIndex        =   13
      Top             =   2880
      Width           =   2175
   End
   Begin VB.CommandButton b_back 
      BackColor       =   &H00FF8080&
      Caption         =   "返回"
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
      Left            =   7440
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   3240
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton daoru 
      BackColor       =   &H00FF8080&
      Caption         =   "保存"
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
      Left            =   5400
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   10
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton b_clear 
      BackColor       =   &H00FF8080&
      Caption         =   "清空"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2760
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   9
      Top             =   4920
      Width           =   1575
   End
   Begin VB.ComboBox Combo_Small1 
      Height          =   300
      Left            =   3840
      TabIndex        =   8
      Top             =   4320
      Visible         =   0   'False
      Width           =   2175
   End
   Begin VB.ComboBox Combo_Version1 
      Height          =   300
      Left            =   3840
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   3480
      Width           =   2175
   End
   Begin VB.TextBox NSR_MC 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Left            =   3840
      TabIndex        =   6
      Top             =   2160
      Width           =   2175
   End
   Begin VB.TextBox NSR_BM 
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      Left            =   3840
      TabIndex        =   5
      Top             =   1560
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "提示:不能从EXCEL直接复制粘贴"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   6240
      TabIndex        =   14
      Top             =   1680
      Width           =   3015
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "所属  行业："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2040
      TabIndex        =   12
      Top             =   2880
      Width           =   2055
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "版  本 号："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2040
      TabIndex        =   4
      Top             =   4200
      Visible         =   0   'False
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "报表  版本："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2040
      TabIndex        =   3
      Top             =   3480
      Width           =   1935
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "纳税人名称："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2040
      TabIndex        =   2
      Top             =   2160
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "纳税人编码："
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   495
      Left            =   2040
      TabIndex        =   1
      Top             =   1560
      Width           =   2055
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "纳税人信息录入"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   360
      TabIndex        =   0
      Top             =   360
      Width           =   3255
   End
End
Attribute VB_Name = "imponsrxx_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_back_Click()
imponsrxx_form.Hide
    login_form.Show
End Sub

Private Sub b_clear_Click()
NSR_BM.Text = ""
NSR_MC.Text = ""
End Sub

Private Sub Combo_Small_Change()

End Sub

Private Sub Combo_hy_Click()
 If Me.combo_hy.Text = "工业企业" Then
 hy = "1"
 ElseIf Me.combo_hy.Text = "房地产企业" Then
 hy = "2"
 Else
 hy = "3"
 End If
 
 
 loadVersionCombox
 
 
 
 
End Sub

Private Sub Combo_Version1_Click()

  Dim versionID As String
  Dim sql As String
  Dim versionRs As ADODB.Recordset
  Dim itemCount As Integer
  
  versionID = getVersionID(Me.Combo_Version1.Text)
  'MsgBox versionId
  'Exit Sub
  
  
  sql = "select small_id from t_baobiao where version_id = '" & versionID & "'"
  Set versionRs = cn.Execute(sql)
  
  itemCount = Me.Combo_Small1.ListCount - 1
  While itemCount >= 0
    Me.Combo_Small1.RemoveItem itemCount
    itemCount = itemCount - 1
  Wend
  
  While Not versionRs.EOF '
     Me.Combo_Small1.AddItem versionRs("small_id")

    versionRs.MoveNext
  Wend
End Sub






Private Sub Combo_Version_Change()

End Sub

Private Sub daoru_Click()
Dim userArray '用户信息数组

Dim nsrbmRs As ADODB.Recordset  '用户的数据库的结果集
  '用户的数据库的结果集
Dim sql As String
Dim userInfoArray
Dim nsrbm As String: nsrbm = ""  '纳税人编码
Dim NSRMC As String: NSRMC = "" '纳税人名称
Dim version As String
'Dim s_version As String
Dim b_id As String
Dim bbid As String
Dim yeno As Boolean
If Len(NSR_BM.Text) < 12 Or Len(NSR_BM.Text) > 18 Then
MsgBox ("纳税人编码长度要求在12位到18位之间")
Exit Sub
Else

If NSR_BM.Text <> "" And NSR_MC.Text <> "" And Combo_Version1.Text <> "" Then
   nsrbm = NSR_BM.Text
   
   nsrbm = Replace(nsrbm, vbCrLf, "")
  
   nsrqc = NSR_MC.Text
   version = Combo_Version1.Text
 ' s_version = Combo_Small1.Text
  Dim rs As ADODB.Recordset
  sql = "select * from t_year_dm where year = '" & version & "'"
  
  Call check_condatabase
  
  Set rs = cn.Execute(sql)
    bbid = rs("version_id")
    rs.Close
    Dim aRs As ADODB.Recordset
  sql = "select * from t_baobiao where version_id = '" & bbid & "' and baobiao_zl= '" & hy & "' "
  
  Call check_condatabase
  
  Set aRs = cn.Execute(sql)
   b_id = aRs("id")
    

Else
   MsgBox ("请输入完整信息")
   Exit Sub
End If
'End If              '数据库操作
Call check_condatabase

sql = "select nsrbm from t_nsrxx where nsrbm = '" & nsrbm & "' and username='" & username & "'"
Set nsrbmRs = cn.Execute(sql)
If Not nsrbmRs.EOF Then
   MsgBox "此纳税人信息已经导入！"
   Exit Sub
End If
nsrbmRs.Close
Set nsrbmRs = Nothing
Dim nsrxxRs As ADODB.Recordset
sql = "select * from t_nsrxx"
Set nsrxxRs = New ADODB.Recordset
Set nsrxxRs.ActiveConnection = cn
nsrxxRs.LockType = adLockOptimistic
nsrxxRs.CursorType = adOpenKeyset

nsrxxRs.Open sql
nsrxxRs.AddNew '添加纳税人信息
nsrxxRs("nsrbm") = nsrbm
nsrxxRs("nsrmc") = nsrqc
nsrxxRs("username") = username
nsrxxRs("zchy") = hy
nsrxxRs.Update
nsrxxRs.Close
Set nsrxxRs = Nothing
Dim bbbbRs As ADODB.Recordset
sql = "select * from t_baobiao_version"
Set bbbbRs = New ADODB.Recordset
Set bbbbRs.ActiveConnection = cn
bbbbRs.LockType = adLockOptimistic
bbbbRs.CursorType = adOpenKeyset

bbbbRs.Open sql
bbbbRs.AddNew
bbbbRs("nsrbm") = nsrbm
bbbbRs("user_name") = username
bbbbRs("baobiao_id") = b_id
bbbbRs.Update
bbbbRs.Close
MsgBox "录入成功！"
yeno = True

Unload Me
Unload main_form
MainForm.Show
main_form.Show
End If
End Sub


Public Sub loadVersionCombox()
    Me.Combo_Version1.Clear
    'Me.Combo_Small.Clear
    
    Dim versionRs As ADODB.Recordset '保存纳税人的结果集
    Dim sql As String
     
   
    
    
    
     
    Call check_condatabase
    sql = "select t_year_dm.year from t_year_dm,t_baobiao where t_baobiao.version_id = t_year_dm.version_id  and t_baobiao.baobiao_zl= '" & hy & " '     "
    
  
    Set versionRs = cn.Execute(sql)
    While Not versionRs.EOF
        If Trim(versionRs("year")) <> "" Then
       ' AddItem nsrRs("nsrbm
        Me.Combo_Version1.AddItem versionRs("year")
        End If
        versionRs.MoveNext
    Wend
    
   ' If Combo_Version.ListCount > 0 Then
   ' Combo_Version.ListIndex = 0
   ' End If
End Sub

Private Sub Form_Load()
  'Me.Width = ScaleX(1024, vbPixels, vbTwips)   '设定窗体的宽度为800像素
  'Me.Height = ScaleY(680, vbPixels, vbTwips)  '设定窗体的高度为680像素
   
  loadVersionCombox
  Me.combo_hy.AddItem ("工业企业")
  Me.combo_hy.AddItem ("房地产企业")
  Me.combo_hy.AddItem ("其他企业")
  
   Me.combo_hy.Text = "工业企业"
 
  'loadNsrCombox   '加载纳税人信息
  'loadBaobiao   '加载报表
  'loadDate  '加载年月下拉框
  'loadVersion
  'isAllowEdit True, Me.F1Book1
End Sub


