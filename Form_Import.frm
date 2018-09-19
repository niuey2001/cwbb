VERSION 5.00
Object = "{B0475000-7740-11D1-BDC3-0020AF9F8E6E}#6.1#0"; "TTF16.ocx"
Object = "{0E59F1D2-1FBE-11D0-8FF2-00A0D10038BC}#1.0#0"; "msscript.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form_Export 
   Caption         =   "数据验证"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7275
   Icon            =   "Form_Import.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8295
   ScaleWidth      =   7275
   Begin VB.TextBox text_warning_mes 
      ForeColor       =   &H00FF0000&
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   10
      Top             =   6480
      Width           =   7095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "导出校验信息"
      Height          =   375
      Left            =   3360
      TabIndex        =   9
      Top             =   3000
      Width           =   1575
   End
   Begin VB.CommandButton Command1 
      Caption         =   "退出"
      Height          =   375
      Left            =   5760
      TabIndex        =   8
      Top             =   3600
      Width           =   1215
   End
   Begin VB.ListBox date_list 
      BackColor       =   &H00C0FFFF&
      Height          =   2580
      Left            =   120
      MultiSelect     =   2  'Extended
      TabIndex        =   1
      Top             =   1560
      Width           =   2895
   End
   Begin TTF160Ctl.F1Book F1Book1 
      Height          =   1215
      Left            =   720
      TabIndex        =   7
      Top             =   2040
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   2143
      _0              =   $"Form_Import.frx":16AC2
      _1              =   $"Form_Import.frx":16ECB
      _2              =   $"Form_Import.frx":172D4
      _3              =   $"Form_Import.frx":176DD
      _4              =   $"Form_Import.frx":17AE6
      _count          =   5
      _ver            =   2
   End
   Begin VB.TextBox text_wrong_mes 
      ForeColor       =   &H000000FF&
      Height          =   1695
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   5
      Top             =   4440
      Width           =   7095
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl1 
      Left            =   5040
      Top             =   1800
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.CommandButton Cmd_Export 
      Caption         =   "验证数据"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   2160
      Width           =   1575
   End
   Begin MSComDlg.CommonDialog CommonDialog_Export 
      Left            =   4200
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSScriptControlCtl.ScriptControl ScriptControl2 
      Left            =   5160
      Top             =   2520
      _ExtentX        =   1005
      _ExtentY        =   1005
   End
   Begin VB.Label Label2 
      Caption         =   "提示性(确认无误后可导出)校验信息："
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   6240
      Width           =   3135
   End
   Begin VB.Label Label3 
      Caption         =   "强制性(必须通过) 校验信息："
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   4200
      Width           =   2535
   End
   Begin VB.Label Label_his_nsr 
      Caption         =   "纳税人编码："
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "选择所属期："
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   3
      Top             =   1200
      Width           =   1935
   End
   Begin VB.Label nsrbm_value 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   1320
      TabIndex        =   2
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "Form_Export"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub nsrmc_valeu_Click()

End Sub
Private Sub Cmd_Export_Click()

Dim dateYear As String
Dim dateSeason As String
Dim ver_name As String
Dim versionStr As String
Dim versionID As String

Dim allValueStr As String


Dim bb_content_id As String
Dim sheetName As String
Dim rs As ADODB.Recordset

Me.text_warning_mes.Text = ""

  
  Dim dateStr As String  '所属期
  dateStr = date_list.Text '200901-200902
  Dim nsrbm As String
  nsrbm = Me.nsrbm_value.Caption

   If dateStr <> "" Then
      dateYear = Mid(dateStr, 1, 4)
      dateSeason = Mid(dateStr, 6, 3)
      dateSeason = date_change(dateSeason)
     
      versionID = get_version(nsrbm, dateYear, dateSeason)
   
      baobiaoName = get_baobiao_name(versionID)
      MsgBox baobiaoName
      ver_name = getVersionNameById(versionID)
      
     main_form.Label_Bb_Value.Caption = ver_name
     loadbaobiao1 (ver_name)
        Call check_condatabase
        sql = "select id,baobiao_name from t_baobiao_content where user_name = '" & username & "' and nsrbm = '" & nsrbm & "' and date_year = '" & dateYear & "' and date_season = '" & dateSeason & "'"
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
        
        
    
   
      
      dateSeason = date_change(dateSeason)
      If validate_exp_data(dateYear, dateSeason, nsrbm, text_wrong_mes, Me.ScriptControl1, main_form.F1Book1) Then
       
         MsgBox "强制性验证通过！"
         isValidateFailed = True
         
      
      Else
        MsgBox "数据输入有误！请查看校验信息"
        isValidateFailed = False
        
      End If
      If validate_ts_data(dateYear, dateSeason, nsrbm, text_wrong_mes, Me.ScriptControl2, main_form.F1Book1) Then
      
      Else
      End If
   Else
   MsgBox "请选择所属期！"
      
   End If

End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
nsrbm = Me.nsrbm_value.Caption
fileName = nsrbm & ".txt"
      
        Set fileObj = CreateObject("Scripting.FileSystemObject")
        Set writeobj = fileObj.CreateTextFile(App.Path & "\" & fileName, True)
        writeobj.writeline ("强制性验证信息：")
        writeobj.writeline (Me.text_wrong_mes.Text)
        writeobj.writeline ("提示性验证信息：")
        writeobj.writeline (Me.text_warning_mes.Text)
        writeobj.Close
MsgBox "成功导出校验信息"

End Sub

Private Sub date_list_DblClick()


Dim dateYear As String
Dim dateSeason As String
Dim ver_name As String
Dim versionStr As String
Dim versionID As String

Dim allValueStr As String


Dim bb_content_id As String
Dim sheetName As String
Dim rs As ADODB.Recordset

Me.text_warning_mes.Text = ""

  
  Dim dateStr As String  '所属期
  dateStr = date_list.Text '200901-200902
  Dim nsrbm As String
  nsrbm = Me.nsrbm_value.Caption

   If dateStr <> "" Then
      dateYear = Mid(dateStr, 1, 4)
      dateSeason = Mid(dateStr, 6, 3)
      dateSeason = date_change(dateSeason)
     
      versionID = get_version(nsrbm, dateYear, dateSeason)
   
      baobiaoName = get_baobiao_name(versionID)
      MsgBox baobiaoName
      ver_name = getVersionNameById(versionID)
     
     main_form.Label_Bb_Value.Caption = ver_name
     loadbaobiao1 (ver_name)
        Call check_condatabase
        sql = "select id,baobiao_name from t_baobiao_content where user_name = '" & username & "' and nsrbm = '" & nsrbm & "' and date_year = '" & dateYear & "' and date_season = '" & dateSeason & "'"
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
        
        
    
  
      
      dateSeason = date_change(dateSeason)
      If validate_exp_data(dateYear, dateSeason, nsrbm, text_wrong_mes, Me.ScriptControl1, main_form.F1Book1) Then
       
         MsgBox "数据验证通过！"
         isValidateFailed = True
         
      
      Else
        MsgBox "数据输入有误！请查看校验信息"
        isValidateFailed = False
        
      End If
   If validate_ts_data(dateYear, dateSeason, nsrbm, text_wrong_mes, Me.ScriptControl2, main_form.F1Book1) Then
      
      Else
      End If
   Else
   MsgBox "请选择所属期！"
      
   End If

End Sub
Private Sub Form_Load()
Me.nsrbm_value.Caption = main_form.Combo_Nsrbm.Text
loadDateList
ScriptControl1.AddObject "textWrongMes", text_wrong_mes '将对象添加进ScriptControl1
ScriptControl2.AddObject "textWrongMes", text_warning_mes
Me.text_wrong_mes.Text = ""
Me.text_warning_mes.Text = ""
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
sql = "select * from t_baobiao_content where user_name = '" & username & "' and nsrbm = '" & Me.nsrbm_value.Caption & "' "
'MsgBox sql
Set dateRs = cn.Execute(sql)
While Not dateRs.EOF
    itemFlag = True
    If Trim(dateRs("date_year")) <> "" And Trim(dateRs("date_season")) <> "" Then
         betweenDate = dateRs("date_year") & "年" & change_date(dateRs("date_season")) & "月"
         'MsgBox betweenDate
         itemCount = Me.date_list.ListCount - 1
         '去除listview中重复元素
         While itemCount >= 0
            'MsgBox historyBaobiao_form.date_list.List(itemCount) & "" & betweenDate
            If Me.date_list.List(itemCount) = betweenDate Then
                  itemFlag = False
            End If
            itemCount = itemCount - 1
         Wend
         If itemFlag Then
            Me.date_list.AddItem betweenDate
         End If
         
    
    End If
    dateRs.MoveNext
Wend
End Sub


