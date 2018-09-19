VERSION 5.00
Object = "{B0475000-7740-11D1-BDC3-0020AF9F8E6E}#6.1#0"; "TTF16.ocx"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form Export_Many_Form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   7365
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7485
   Icon            =   "Export_Many_Form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   7485
   Begin MSComctlLib.ListView ListView1 
      Height          =   3495
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   6165
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
   Begin TTF160Ctl.F1Book F1Book1 
      Height          =   1695
      Left            =   600
      TabIndex        =   13
      Top             =   1560
      Width           =   1815
      _ExtentX        =   3201
      _ExtentY        =   2990
      _0              =   $"Export_Many_Form.frx":16AC2
      _1              =   $"Export_Many_Form.frx":16ECB
      _2              =   $"Export_Many_Form.frx":172D4
      _3              =   $"Export_Many_Form.frx":176DD
      _4              =   $"Export_Many_Form.frx":17AE6
      _count          =   5
      _ver            =   2
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   6000
      Top             =   3120
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.TextBox text_wrong_mes 
      ForeColor       =   &H000000FF&
      Height          =   2295
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   11
      Top             =   5040
      Width           =   7455
   End
   Begin VB.ComboBox Combo_YearSeason 
      Height          =   300
      Left            =   3960
      Style           =   2  'Dropdown List
      TabIndex        =   10
      Top             =   2520
      Width           =   3255
   End
   Begin VB.CommandButton b_quit 
      Caption         =   "退出"
      Height          =   495
      Left            =   5880
      TabIndex        =   8
      Top             =   3960
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Export 
      Caption         =   "导出"
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   3960
      Width           =   1215
   End
   Begin VB.TextBox text_import_path 
      Enabled         =   0   'False
      Height          =   375
      Left            =   3960
      TabIndex        =   6
      Top             =   1440
      Width           =   3255
   End
   Begin VB.CommandButton Cmd_Choose 
      Caption         =   "选择路径.."
      Height          =   375
      Left            =   6000
      TabIndex        =   5
      Top             =   960
      Width           =   1215
   End
   Begin VB.CommandButton cmd_choose_all 
      Caption         =   "全选"
      Height          =   375
      Left            =   1800
      TabIndex        =   1
      Top             =   480
      Width           =   855
   End
   Begin VB.CommandButton cmd_cancle_choose 
      Caption         =   "取消全选"
      Height          =   375
      Left            =   2760
      TabIndex        =   0
      Top             =   480
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialog_Export 
      Left            =   6840
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label4 
      Caption         =   "校验信息："
      Height          =   255
      Left            =   240
      TabIndex        =   12
      Top             =   4800
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "导出至："
      Height          =   255
      Left            =   3960
      TabIndex        =   9
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label2 
      Caption         =   "纳税人列表："
      ForeColor       =   &H00FF0000&
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   600
      Width           =   1575
   End
   Begin VB.Label Label3 
      Caption         =   "导出指标数据"
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
      Left            =   3120
      TabIndex        =   3
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "Export_Many_Form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_quit_Click()
Unload Me

End Sub

Private Sub cmd_cancle_choose_Click()
For i = 1 To ListView1.ListItems.Count
       ListView1.ListItems(i).Checked = False '''true打勾,false取消打勾
Next
End Sub

Private Sub cmd_choose_all_Click()
For i = 1 To ListView1.ListItems.Count
       ListView1.ListItems(i).Checked = True '''true打勾,false取消打勾
Next
End Sub

Private Sub Cmd_Choose_Click()
Dim filePath As String

filePath = GetFolder("打开一个目录", Me.CommonDialog_Export)
Me.text_import_path.Text = filePath
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
    sql = "select nsrbm,nsrmc from t_nsrxx where username='" & username & "'"
    Set nsrRs = cn.Execute(sql)
    While Not nsrRs.EOF
        If Trim(nsrRs("nsrbm")) <> "" Then
            
        Set lv = ListView1.ListItems.Add(, , nsrRs("nsrbm"))
            lv.SubItems(1) = nsrRs("nsrmc")
        End If
        nsrRs.MoveNext
    Wend
    loadDateList
    
    ScriptControl1.AddObject "textWrongMes", text_wrong_mes  '将对象添加进ScriptControl1
    Me.text_wrong_mes.Text = ""
End Sub

Private Sub loadDateList()

Dim dateRs As ADODB.Recordset  '保存报表所属期的结果集
Dim sql As String
Dim version As String
'version = Me.lable_version.Caption
Dim betweenDate As String
Dim itemCount As Integer
Dim itemFlag As Boolean
Call check_condatabase
sql = "select * from t_baobiao_content where user_name = '" & username & "'"
'MsgBox sql
Set dateRs = cn.Execute(sql)
While Not dateRs.EOF
    itemFlag = True
    If Trim(dateRs("date_year")) <> "" And Trim(dateRs("date_season")) <> "" Then
         betweenDate = dateRs("date_year") & "年" & dateRs("date_season") & "季度"
         'MsgBox betweenDate
         itemCount = Me.Combo_YearSeason.ListCount - 1
         '去除listview中重复元素
         While itemCount >= 0
            'MsgBox historyBaobiao_form.date_list.List(itemCount) & "" & betweenDate
            If Me.Combo_YearSeason.List(itemCount) = betweenDate Then
                  itemFlag = False
            End If
            itemCount = itemCount - 1
         Wend
         If itemFlag Then
          Me.Combo_YearSeason.AddItem betweenDate
         End If
         
    
    End If
    dateRs.MoveNext
Wend
End Sub

