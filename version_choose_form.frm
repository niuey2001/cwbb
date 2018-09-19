VERSION 5.00
Begin VB.Form version_choose_form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "报表版本选择窗口"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6360
   Icon            =   "version_choose_form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   6360
   Begin VB.CommandButton Cmd_Quit 
      Caption         =   "退出"
      Height          =   495
      Left            =   4560
      TabIndex        =   11
      Top             =   5040
      Width           =   1455
   End
   Begin VB.CommandButton Cmd_Version_Save 
      Caption         =   "保存"
      Height          =   495
      Left            =   2880
      TabIndex        =   10
      Top             =   5040
      Width           =   1455
   End
   Begin VB.Frame Frame_version 
      Caption         =   "报表版本设置"
      Height          =   2655
      Left            =   240
      TabIndex        =   5
      Top             =   2160
      Width           =   5775
      Begin VB.ComboBox Combo_Small 
         BackColor       =   &H00C0FFFF&
         Height          =   300
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   0
         Visible         =   0   'False
         Width           =   2415
      End
      Begin VB.ComboBox Combo_hy 
         BackColor       =   &H00C0FFFF&
         Height          =   300
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   600
         Width           =   2415
      End
      Begin VB.ComboBox Combo_Version 
         BackColor       =   &H00C0FFFF&
         Height          =   300
         ItemData        =   "version_choose_form.frx":16AC2
         Left            =   1800
         List            =   "version_choose_form.frx":16AC4
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1200
         Width           =   2415
      End
      Begin VB.Label Label6 
         Caption         =   "提示：企业采用2006年公布的财务会计准则核算的，适用新版(2007年版),否则适用旧版(2005年版)"
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   12
         Top             =   1920
         Width           =   5295
      End
      Begin VB.Label Label1 
         Caption         =   "行    业："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   600
         Width           =   1335
      End
      Begin VB.Label Label_Version 
         Caption         =   "报表版本："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   6
         Top             =   1200
         Width           =   1575
      End
   End
   Begin VB.Frame frame_nsrxx 
      Caption         =   "纳税人信息"
      Height          =   1575
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5775
      Begin VB.Label Label_Nsrmc 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1920
         TabIndex        =   4
         Top             =   960
         Width           =   3255
      End
      Begin VB.Label Label_Nsrbm 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1920
         TabIndex        =   3
         Top             =   360
         Width           =   3255
      End
      Begin VB.Label Label3 
         Caption         =   "纳税人编码："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   2
         Top             =   360
         Width           =   1575
      End
      Begin VB.Label Label2 
         Caption         =   "纳税人名称："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   600
         TabIndex        =   1
         Top             =   960
         Width           =   1575
      End
   End
End
Attribute VB_Name = "version_choose_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Frame1_DragDrop(Source As Control, X As Single, Y As Single)

End Sub

Private Sub Cmd_Quit_Click()
Unload Me
End Sub

Private Sub Cmd_Version_Save_Click()
    
    If Me.Combo_Version.Text <> "" And Me.combo_hy.Text <> "" Then
    
    Dim hy As String
    
     If Me.combo_hy.Text = "工业企业" Then
 hy = "1"
 ElseIf Me.combo_hy.Text = "房地产企业" Then
 hy = "2"
 Else
 hy = "3"
 End If
 
 
    
    
    ' Me.Combo_Version.Text  me.Combo_Small.Text  Me.Label_Nsrbm.Caption
    saveNsrVersion Me.Combo_Version.Text, Me.Label_Nsrbm.Caption, hy
    
    updatensr_hy Me.combo_hy.Text, Me.Label_Nsrbm.Caption
    
    MsgBox "保存成功！"
    Unload Me
    Unload main_form
    main_form.Show
    
    
    Else
        MsgBox "报表版本和行业必须选择"
    End If

    
    
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

Private Sub Combo_Version_Click()
  Dim versionID As String
  Dim sql As String
  Dim versionRs As ADODB.Recordset
  Dim itemCount As Integer
  
  versionID = getVersionID(Me.Combo_Version.Text)
  'MsgBox versionId
  'Exit Sub
  
  
  sql = "select small_id from t_baobiao where version_id = '" & versionID & "'"
  Set versionRs = cn.Execute(sql)
  
  itemCount = Me.Combo_Small.ListCount - 1
  While itemCount >= 0
    Me.Combo_Small.RemoveItem itemCount
    itemCount = itemCount - 1
  Wend
  
 ' While Not versionRs.EOF '
  '   Me.Combo_Small.AddItem versionRs("small_id")

   ' versionRs.MoveNext
 ' Wend
End Sub

Private Sub Form_Load()
 Me.combo_hy.AddItem ("工业企业")
  Me.combo_hy.AddItem ("房地产企业")
  Me.combo_hy.AddItem ("其他企业")
  
   Me.combo_hy.Text = "工业企业"
loadVersionCombox
End Sub
Public Sub loadVersionCombox()
    Me.Combo_Version.Clear
    Dim versionRs As ADODB.Recordset '保存纳税人的结果集
    Dim sql As String
    
    Call check_condatabase
    sql = "select t_year_dm.year from t_year_dm,t_baobiao where t_baobiao.version_id = t_year_dm.version_id and baobiao_zl='" & hy & " '"
    Set versionRs = cn.Execute(sql)
    While Not versionRs.EOF
        If Trim(versionRs("year")) <> "" Then
       ' AddItem nsrRs("nsrbm
        Me.Combo_Version.AddItem versionRs("year")
        End If
        versionRs.MoveNext
    Wend
    
   ' If Combo_Version.ListCount > 0 Then
   ' Combo_Version.ListIndex = 0
   ' End If
End Sub

