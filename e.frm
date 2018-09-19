VERSION 5.00
Begin VB.Form baobiao_option_form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   3705
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6555
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   3705
   ScaleWidth      =   6555
   Begin VB.CommandButton cmd_quit 
      Caption         =   "退出"
      Height          =   495
      Left            =   5280
      TabIndex        =   1
      Top             =   2880
      Width           =   1095
   End
   Begin VB.CommandButton cmd_del 
      Caption         =   "删除"
      Height          =   495
      Left            =   5280
      TabIndex        =   0
      Top             =   2160
      Width           =   1095
   End
   Begin VB.PictureBox ListView1 
      BackColor       =   &H00C0FFFF&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   120
      ScaleHeight     =   2835
      ScaleWidth      =   4875
      TabIndex        =   3
      Top             =   720
      Width           =   4935
   End
   Begin VB.Label Label1 
      Caption         =   "报表模板列表："
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
      Left            =   360
      TabIndex        =   2
      Top             =   240
      Width           =   3615
   End
End
Attribute VB_Name = "baobiao_option_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Form_Load()
 With ListView1
        .ColumnHeaders.Clear
          .ListItems.Clear
        .ColumnHeaders.Add , , "报表名称"
        '设置“卷”的显示宽度
        .ColumnHeaders(1).Width = 1900

    End With
    
    
    Dim baobiaoRs As ADODB.Recordset '保存纳税人的结果集
    Dim sql As String
    
    Call check_condatabase
    sql = "select id,baobiao_name from t_baobiao"
    Set baobiaoRs = cn.Execute(sql)
    While Not baobiaoRs.EOF
        If Trim(baobiaoRs("id")) <> "" Then
            
        Set lv = ListView1.ListItems.Add(, , baobiaoRs("baobiao_name"))
        End If
        baobiaoRs.MoveNext
    Wend
End Sub

