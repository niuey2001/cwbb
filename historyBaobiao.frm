VERSION 5.00
Begin VB.Form historyBaobiao_form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "报表历史数据"
   ClientHeight    =   6420
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6795
   Icon            =   "historyBaobiao.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6420
   ScaleWidth      =   6795
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Frame1 
      Height          =   3495
      Left            =   3240
      TabIndex        =   8
      Top             =   1080
      Width           =   3375
      Begin VB.Label l_small_id 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1080
         TabIndex        =   16
         Top             =   840
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label Label6 
         Caption         =   "版 本 号："
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   960
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.Label l_baobiaoNames 
         BackColor       =   &H00C0FFFF&
         Height          =   1215
         Left            =   1080
         TabIndex        =   14
         Top             =   2040
         Width           =   1935
      End
      Begin VB.Label Label4 
         Caption         =   "录入报表："
         Height          =   375
         Left            =   120
         TabIndex        =   13
         Top             =   2160
         Width           =   975
      End
      Begin VB.Label l_time_value 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1080
         TabIndex        =   12
         Top             =   1440
         Width           =   2175
      End
      Begin VB.Label l_version_value 
         BackColor       =   &H00C0FFFF&
         Height          =   375
         Left            =   1080
         TabIndex        =   11
         Top             =   600
         Width           =   1575
      End
      Begin VB.Label Label3 
         Caption         =   "录入时间："
         Height          =   255
         Left            =   120
         TabIndex        =   10
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "版    本："
         Height          =   375
         Left            =   120
         TabIndex        =   9
         Top             =   600
         Width           =   1215
      End
   End
   Begin VB.CommandButton b_quit 
      Caption         =   "退出"
      Height          =   495
      Left            =   3720
      TabIndex        =   3
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton view 
      Caption         =   "查看"
      Height          =   495
      Left            =   3720
      TabIndex        =   2
      Top             =   4680
      Width           =   1455
   End
   Begin VB.ListBox date_list 
      BackColor       =   &H00C0FFFF&
      Height          =   4920
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   3015
   End
   Begin VB.Label nsrmc_valeu 
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   235
      Width           =   1935
   End
   Begin VB.Label nsrbm_value 
      Height          =   255
      Left            =   1320
      TabIndex        =   6
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label_Nsrmc 
      Caption         =   "纳税人名称："
      Height          =   255
      Left            =   3240
      TabIndex        =   5
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label_his_nsr 
      Caption         =   "纳税人编码："
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   240
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "所属期："
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   1935
   End
End
Attribute VB_Name = "historyBaobiao_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_quit_Click()
Unload Me

End Sub

Private Sub Label5_Click()

End Sub

Private Sub date_list_Click()
Dim baobiaoName As String: baobiaoName = ""
Dim versionID As String: versionID = ""
Dim versionName As String: versionName = ""
Dim smallVersionID As String: smallVersionID = ""
Dim createTime As String: createTime = ""
Dim baobiaoNamesStr As String: baobiaoNamesStr = ""
Dim dateSeason As String

dateStr = date_list.Text  '200901-200902
If Trim(dateStr) <> "" Then
    dateYear = Mid(dateStr, 1, 4)
    dateSeason = Mid(dateStr, 6, 3)
       ' MsgBox dateYear & "  " & dateSeason
        dateSeason = date_change(dateSeason)
   
    Call check_condatabase
    sql = "select baobiao_name,version,create_time from t_baobiao_content where user_name = '" & username & "' and nsrbm = '" & Me.nsrbm_value.Caption & "' and date_year = '" & dateYear & "' and date_season = '" & dateSeason & "'"
    Set rs = cn.Execute(sql)
    
    While Not rs.EOF
        If versionID = "" Then
            versionID = rs("version")
        End If
       ' If smallVersionID = "" Then
        '    smallVersionID = rs("small_version_id")
        'End If
        If createTime = "" Then
            createTime = rs("create_time")
        End If
        
        baobiaoName = rs("baobiao_name")
        baobiaoNamesStr = baobiaoNamesStr + baobiaoName & vbCrLf
        rs.MoveNext
    Wend
    rs.Close
    Set rs = Nothing
   versionName = getVersionNameById(versionID)
   Me.l_version_value = versionName
   Me.l_time_value.Caption = createTime
   'Me.l_small_id = smallVersionID
   Me.l_baobiaoNames = baobiaoNamesStr
   
  'Public Function showData4NameSheet(sheetName As String, showDataStr As String, F1Book1 As F1Book)
'    If bb_content_id <> "0" Then
'           allValueStr = operate_form.getValuesById(bb_content_id)
'           operate_form.showData (allValueStr)
'    Else
'        MsgBox "此报表没有保存！"
'    End If

End If
End Sub
Private Sub date_list_DblClick()
Dim allValueStr As String
Dim dateStr As String
Dim dateYear As String
Dim dateSeason As String
Dim bb_content_id As String
Dim sheetName As String
Dim rs As ADODB.Recordset


dateStr = date_list.Text  '200901-200902
If Trim(dateStr) <> "" Then
 
     main_form.Label_Bb_Value.Caption = Me.l_version_value.Caption
     loadbaobiao1 (Me.l_version_value.Caption)
          
        
         dateYear = Mid(dateStr, 1, 4)
        dateSeason = Mid(dateStr, 6, 3)
       ' MsgBox dateYear & "  " & dateSeason
        dateSeason = date_change(dateSeason)
        Call check_condatabase
        sql = "select id,baobiao_name from t_baobiao_content where user_name = '" & username & "' and nsrbm = '" & Me.nsrbm_value.Caption & "' and date_year = '" & dateYear & "' and date_season = '" & dateSeason & "'"
        Set rs = cn.Execute(sql)
        
        While Not rs.EOF
            sheetName = rs("baobiao_name")
            bb_content_id = rs("id")
            'MsgBox sheetName & "    " & bb_content_id
            allValueStr = getValuesById(bb_content_id)
            'MsgBox allValueStr
            showData4NameSheet sheetName, allValueStr, main_form.F1Book1
            
            rs.MoveNext
        Wend
        rs.Close
        Set rs = Nothing
        dateSeason = change_date(dateSeason)
        main_form.CB_Year.Text = dateYear
        main_form.CB_Season.Text = dateSeason
        main_form.F1Book1.Sheet = 1
        If hy = "1" Or hy = "2" Then
        
           main_form.F1Book1.ObjValue(pid3) = main_form.F1Book1.EntryRC(3, 3)
           Else
           
           End If
          main_form.F1Book1.ObjValue(pid) = main_form.F1Book1.EntryRC(9, 3)
           main_form.F1Book1.ObjValue(pid2) = main_form.F1Book1.EntryRC(10, 3)
        
        
        Unload Me
  
  
    
Else
    MsgBox "请选择正确所属期！"
End If
End Sub

Private Sub view_Click()
Dim allValueStr As String
Dim dateStr As String
Dim dateYear As String
Dim dateSeason As String
Dim bb_content_id As String
Dim sheetName As String
Dim rs As ADODB.Recordset



dateStr = date_list.Text  '200901-200902
If Trim(dateStr) <> "" Then

     main_form.Label_Bb_Value.Caption = Me.l_version_value.Caption
     loadbaobiao1 (Me.l_version_value.Caption)
          
        
         dateYear = Mid(dateStr, 1, 4)
        dateSeason = Mid(dateStr, 6, 3)
       ' MsgBox dateYear & "  " & dateSeason
        dateSeason = date_change(dateSeason)
        Call check_condatabase
        sql = "select id,baobiao_name from t_baobiao_content where user_name = '" & username & "' and nsrbm = '" & Me.nsrbm_value.Caption & "' and date_year = '" & dateYear & "' and date_season = '" & dateSeason & "'"
        Set rs = cn.Execute(sql)
        
        While Not rs.EOF
            sheetName = rs("baobiao_name")
            bb_content_id = rs("id")
            'MsgBox sheetName & "    " & bb_content_id
            allValueStr = getValuesById(bb_content_id)
            'MsgBox allValueStr
            showData4NameSheet sheetName, allValueStr, main_form.F1Book1
            
            rs.MoveNext
        Wend
        rs.Close
        Set rs = Nothing
        dateSeason = change_date(dateSeason)
        main_form.CB_Year.Text = dateYear
        main_form.CB_Season.Text = dateSeason
        main_form.F1Book1.Sheet = 1
        If hy = "1" Or hy = "2" Then
        
        main_form.F1Book1.ObjValue(pid3) = main_form.F1Book1.EntryRC(3, 3)
        Else
        
        End If
          main_form.F1Book1.ObjValue(pid) = main_form.F1Book1.EntryRC(9, 3)
           main_form.F1Book1.ObjValue(pid2) = main_form.F1Book1.EntryRC(10, 3)
        
        
       Unload Me
    
    
Else
    MsgBox "请选择正确所属期！"
End If

End Sub
