VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Del_history_form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "管理历史数据"
   ClientHeight    =   6840
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8835
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6840
   ScaleWidth      =   8835
   StartUpPosition =   3  '窗口缺省
   Begin VB.ComboBox c_season 
      Height          =   300
      Left            =   7320
      TabIndex        =   12
      Top             =   3480
      Width           =   855
   End
   Begin VB.ComboBox c_year 
      Height          =   300
      Left            =   6000
      TabIndex        =   11
      Top             =   3480
      Width           =   855
   End
   Begin VB.CommandButton c_x 
      Caption         =   "查询"
      Height          =   495
      Left            =   4920
      TabIndex        =   10
      Top             =   5160
      Width           =   855
   End
   Begin VB.CommandButton quite 
      Caption         =   "退出"
      Height          =   495
      Left            =   7440
      TabIndex        =   9
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton del 
      Caption         =   "删除"
      Height          =   495
      Left            =   6120
      TabIndex        =   8
      Top             =   5160
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "取消全选"
      Height          =   375
      Left            =   3120
      TabIndex        =   4
      Top             =   720
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "全选"
      Height          =   375
      Left            =   2040
      TabIndex        =   3
      Top             =   720
      Width           =   855
   End
   Begin MSComctlLib.ListView ListView1 
      Height          =   4575
      Left            =   0
      TabIndex        =   2
      Top             =   1200
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   8070
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      Checkboxes      =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.ComboBox C_nsrbm 
      Height          =   300
      Left            =   6000
      TabIndex        =   0
      Text            =   "Combo1"
      Top             =   1920
      Width           =   2295
   End
   Begin VB.Label Label7 
      Caption         =   $"Del_history_form.frx":0000
      Height          =   495
      Left            =   240
      TabIndex        =   15
      Top             =   6120
      Width           =   7935
   End
   Begin VB.Label Label6 
      Caption         =   "月"
      Height          =   255
      Left            =   8280
      TabIndex        =   14
      Top             =   3480
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "年"
      Height          =   255
      Left            =   6960
      TabIndex        =   13
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "报表所属期："
      Height          =   255
      Left            =   4800
      TabIndex        =   7
      Top             =   3480
      Width           =   1095
   End
   Begin VB.Label Label3 
      Caption         =   "纳税人编码："
      Height          =   375
      Left            =   4800
      TabIndex        =   6
      Top             =   1920
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "管理历史数据"
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
      TabIndex        =   5
      Top             =   120
      Width           =   3615
   End
   Begin VB.Label Label1 
      Caption         =   " 历史数据列表："
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   720
      Width           =   1575
   End
End
Attribute VB_Name = "Del_history_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command5_Click()

End Sub

Private Sub c_x_Click()
 ListView1.ListItems.Clear
      
Dim nsrRs As ADODB.Recordset '保存纳税人的结果集
    Dim sql As String
    Dim year As String
    Dim season As String
    year = Me.c_year.Text
    season = date_change(Me.c_season.Text)
    If Me.C_nsrbm.Text = "" Then
      If season <> "" And year <> "" Then
      sql = "select distinct nsrbm, date_year,date_season from t_baobiao_content where user_name='" & username & "' and date_year='" & year & "'and date_season='" & season & "'"
       ElseIf season = "" And year = "" Then
      'MsgBox ("纳税人编码或者报表日期中选一个指标")
      sql = "select distinct nsrbm, date_year,date_season from t_baobiao_content where user_name='" & username & "' "
      ElseIf season = "" Or year = "" Then
      MsgBox ("输入完整日期")
      sql = ""
  
      End If
    ElseIf Me.C_nsrbm.Text <> "" And season = "" And year = "" Then
    sql = "select distinct nsrbm, date_year,date_season from t_baobiao_content where user_name='" & username & "' and nsrbm='" & Me.C_nsrbm.Text & "'"
    ElseIf Me.C_nsrbm.Text <> "" And season <> "" And year <> "" Then
    sql = "select distinct nsrbm, date_year,date_season from t_baobiao_content where user_name='" & username & "' and nsrbm='" & Me.C_nsrbm.Text & "'and date_year='" & year & "'and date_season='" & season & "'"
    Else
    MsgBox ("输入完整日期")
    sql = ""
    End If
    Call check_condatabase
 ' MsgBox (sql)
 If sql <> "" Then

    Set nsrRs = cn.Execute(sql)
    While Not nsrRs.EOF
        If Trim(nsrRs("nsrbm")) <> "" Then
            
        Set lv = ListView1.ListItems.Add(, , nsrRs("nsrbm"))
            lv.SubItems(1) = nsrRs("date_year")
        
            lv.SubItems(2) = change_date(nsrRs("date_season"))
           ' lv.SubItems(3) = nsrRs("baobiao_name")
           
        End If
        nsrRs.MoveNext
    Wend

End If
End Sub

Private Sub Command1_Click()
For i = 1 To ListView1.ListItems.Count
       ListView1.ListItems(i).Checked = True '''true打勾,false取消打勾
Next
End Sub

Private Sub Command2_Click()
For i = 1 To ListView1.ListItems.Count
       ListView1.ListItems(i).Checked = False '''true打勾,false取消打勾
Next
End Sub

Private Sub del_Click()
Dim nsrbm As String
Dim year As String
Dim season As String
Dim sql As String
Dim aRs As ADODB.Recordset

        For i = 1 To ListView1.ListItems.Count
               If ListView1.ListItems(i).Checked Then
                  
                   ' If nsrbmStr = "" Then
                        nsrbm = ListView1.ListItems(i).Text
                        year = ListView1.ListItems(i).SubItems(1)
                        season = ListView1.ListItems(i).SubItems(2)
               Else
               End If
        Call check_condatabase
        sql = "select id,baobiao_name from t_baobiao_content where nsrbm='" & nsrbm & "' and date_year='" & year & "'and date_season='" & date_change(season) & "'and user_name='" & username & "'"
        Set aRs = cn.Execute(sql)
       While Not aRs.EOF
        sql = "delete from t_baobiao_value where bb_content_id='" & aRs("id") & "' "
        cn.Execute (sql)
        
        sql = "delete from t_baobiao_content where nsrbm='" & nsrbm & "' and date_year='" & year & "'and date_season='" & date_change(season) & "'and user_name='" & username & "' and baobiao_name='" & aRs("baobiao_name") & "'"
       cn.Execute (sql)
       aRs.MoveNext
       Wend
       
       
     Next
     MsgBox ("删除成功")
End Sub

Private Sub Form_Load()
 With ListView1
        .ColumnHeaders.Clear
          .ListItems.Clear
        .ColumnHeaders.Add , , "纳税人编码"
        '设置“卷”的显示宽度
        .ColumnHeaders(1).Width = 1500
        .ColumnHeaders.Add , , "报表所属年"
         .ColumnHeaders(2).Width = 1400
          .ColumnHeaders.Add , , "报表所属月"
         .ColumnHeaders(3).Width = 1400
         ' .ColumnHeaders.Add , , "报表类别"
         '.ColumnHeaders(4).Width = 1600
         .ColumnHeaders(3).Alignment = lvwColumnLeft
    End With
LoadNsrbm_Combox '加载纳税人信息
LoadDate_Combox '加载年月下拉框

End Sub

Public Sub LoadDate_Combox()
   Me.c_year.Clear
   Me.c_season.Clear
   Dim i As Integer
  For i = 2009 To 2030
    Me.c_year.AddItem CStr(i)
  Next i
  
    Me.c_season.AddItem "1-3"
    Me.c_season.AddItem "1-6"
    Me.c_season.AddItem "1-9"
    Me.c_season.AddItem "1-12"
      
End Sub
Public Sub LoadNsrbm_Combox()

    C_nsrbm.Clear
    Dim nsrRs As ADODB.Recordset '保存纳税人的结果集
    Dim sql As String
    
    Call check_condatabase
    sql = "select nsrbm from t_nsrxx where username='" & username & "'"
    Set nsrRs = cn.Execute(sql)
    While Not nsrRs.EOF
        If Trim(nsrRs("nsrbm")) <> "" Then
        Me.C_nsrbm.AddItem nsrRs("nsrbm")
        End If
        nsrRs.MoveNext
    Wend
    
    'If C_nsrbm.ListCount > 0 Then
    'Me.C_nsrbm.ListIndex = 0
    'End If
End Sub

Private Sub quite_Click()
Unload Me
End Sub
