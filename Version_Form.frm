VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Version_Form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5655
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9090
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   9090
   Begin VB.CommandButton cmd_cancle_choose 
      Caption         =   "取消全选"
      Height          =   375
      Left            =   2880
      TabIndex        =   13
      Top             =   600
      Width           =   975
   End
   Begin VB.CommandButton cmd_choose_all 
      Caption         =   "全选"
      Height          =   375
      Left            =   1920
      TabIndex        =   12
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   735
      Left            =   6840
      TabIndex        =   11
      Top             =   4320
      Width           =   1575
   End
   Begin VB.CommandButton cmd_version_choose 
      Caption         =   "确定"
      Height          =   735
      Left            =   4560
      TabIndex        =   10
      Top             =   4320
      Width           =   1575
   End
   Begin VB.Frame Frame_version 
      Caption         =   "报表版本设置"
      Height          =   2655
      Left            =   4200
      TabIndex        =   1
      Top             =   840
      Width           =   4575
      Begin VB.ComboBox Combo_Version 
         BackColor       =   &H00C0FFFF&
         Height          =   300
         ItemData        =   "Version_Form.frx":0000
         Left            =   1800
         List            =   "Version_Form.frx":0002
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   600
         Width           =   2415
      End
      Begin VB.ComboBox Combo_Small 
         BackColor       =   &H00C0FFFF&
         Height          =   300
         Left            =   1800
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1200
         Width           =   2415
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
         TabIndex        =   7
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label1 
         Caption         =   "版 本 号："
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
         Width           =   1335
      End
      Begin VB.Label Label6 
         Caption         =   "提示：企业采用2006年公布的财务会计准则核算的，适用新版(2007年版),否则适用旧版(2005年版)"
         ForeColor       =   &H000000FF&
         Height          =   495
         Left            =   120
         TabIndex        =   5
         Top             =   1920
         Width           =   3975
      End
      Begin VB.Label Label4 
         Caption         =   "版本和版本号可选项若为空请先导入模板"
         ForeColor       =   &H000000FF&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   3495
      End
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4575
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   12648447
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.Label Label3 
      Caption         =   "设置报表版本"
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
      Left            =   3240
      TabIndex        =   9
      Top             =   120
      Width           =   4695
   End
   Begin VB.Label Label2 
      Caption         =   "纳税人列表："
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "Version_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmd_cancle_choose_Click()
For I = 1 To ListView1.ListItems.Count
       ListView1.ListItems(I).Checked = False '''true打勾,false取消打勾
Next
End Sub

Private Sub cmd_choose_all_Click()
For I = 1 To ListView1.ListItems.Count
       ListView1.ListItems(I).Checked = True '''true打勾,false取消打勾
Next
End Sub

Private Sub cmd_version_choose_Click()
Dim NSRBM As String
Dim versionName As String
Dim small_version As String
versionName = Me.Combo_Version.Text
small_version = Me.Combo_Small.Text
If versionName <> "" And small_version <> "" Then

    For I = 1 To ListView1.ListItems.Count
           If ListView1.ListItems(I).Checked Then
                    NSRBM = ListView1.ListItems(I).Text
                    If NSRBM <> "" Then
                    saveNsrVersion versionName, NSRBM
                    End If
                    
            End If
    Next
MsgBox "保存完毕"
Else
   MsgBox "版本和版本号不可为空！"
End If


End Sub

Private Sub Command2_Click()
Unload Me

End Sub

Private Sub Form_Load()
 With ListView1
        .ColumnHeaders.Clear
          .ListItems.Clear
        .ColumnHeaders.Add , , "纳税人编码"
        '设置“卷”的显示宽度
        .ColumnHeaders(1).Width = 1200
        .ColumnHeaders.Add , , "纳税人名称"
         .ColumnHeaders(1).Width = 1600
         .ColumnHeaders(2).Alignment = lvwColumnLeft
    End With
    
    
    Dim nsrRs As ADODB.Recordset '保存纳税人的结果集
    Dim sql As String
    
    Call check_condatabase
    sql = "select nsrbm,nsrmc from t_nsrxx where username='" & userName & "'"
    Set nsrRs = cn.Execute(sql)
    While Not nsrRs.EOF
        If Trim(nsrRs("nsrbm")) <> "" Then
            
        Set lv = ListView1.ListItems.Add(, , nsrRs("nsrbm"))
            lv.SubItems(1) = nsrRs("nsrmc")
        End If
        nsrRs.MoveNext
    Wend
    
    
    
    
loadVersionCombox
            
            
End Sub
Public Sub loadVersionCombox()
    Me.Combo_Version.Clear
    Dim versionRs As ADODB.Recordset '保存纳税人的结果集
    Dim sql As String
    
    Call check_condatabase
    sql = "select t_year_dm.year from t_year_dm,t_baobiao where t_baobiao.version_id = t_year_dm.version_id"
    Set versionRs = cn.Execute(sql)
    While Not versionRs.EOF
        If Trim(versionRs("year")) <> "" Then
       ' AddItem nsrRs("nsrbm
        Me.Combo_Version.AddItem versionRs("year")
        'me.Combo_Small
        End If
        versionRs.MoveNext
    Wend
    
   ' If Combo_Version.ListCount > 0 Then
   ' Combo_Version.ListIndex = 0
   ' End If
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
  
  While Not versionRs.EOF '
     Me.Combo_Small.AddItem versionRs("small_id")

    versionRs.MoveNext
  Wend
End Sub

