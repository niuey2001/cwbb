VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MainForm 
   BackColor       =   &H8000000C&
   Caption         =   "财务指标采集系统"
   ClientHeight    =   9435
   ClientLeft      =   300
   ClientTop       =   990
   ClientWidth     =   15210
   Icon            =   "MainForm.frx":0000
   LinkTopic       =   "MDIForm1"
   WindowState     =   2  'Maximized
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   435
      Left            =   0
      TabIndex        =   0
      Top             =   9000
      Width           =   15210
      _ExtentX        =   26829
      _ExtentY        =   767
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   6174
            MinWidth        =   6174
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4410
            MinWidth        =   4410
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            TextSave        =   "2013-04-17"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Enabled         =   0   'False
            TextSave        =   "NUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Enabled         =   0   'False
            TextSave        =   "CAPS"
         EndProperty
      EndProperty
   End
   Begin VB.Menu user_manager 
      Caption         =   "用户管理"
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu modify_psw 
         Caption         =   "修改密码"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu change_user 
         Caption         =   "切换用户"
      End
      Begin VB.Menu line5 
         Caption         =   "-"
      End
      Begin VB.Menu quit_sys 
         Caption         =   "退出系统"
      End
   End
   Begin VB.Menu option 
      Caption         =   "模板操作"
      Begin VB.Menu fg33 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu baobiaoDel 
         Caption         =   "管理历史数据"
      End
      Begin VB.Menu fenge87 
         Caption         =   "-"
      End
      Begin VB.Menu history 
         Caption         =   "查看历史报表"
      End
      Begin VB.Menu feng111 
         Caption         =   "-"
      End
      Begin VB.Menu version 
         Caption         =   "报表版本设置"
      End
      Begin VB.Menu line88 
         Caption         =   "-"
      End
      Begin VB.Menu showOperateForm 
         Caption         =   "显示主操作界面"
      End
      Begin VB.Menu line12 
         Caption         =   "-"
      End
      Begin VB.Menu exp 
         Caption         =   "验证数据"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu TJNSR 
         Caption         =   "添加纳税人"
      End
   End
   Begin VB.Menu bf_hf 
      Caption         =   "数据备份与恢复"
      Begin VB.Menu line101 
         Caption         =   "-"
      End
      Begin VB.Menu bf 
         Caption         =   "数据备份"
      End
      Begin VB.Menu line1211 
         Caption         =   "-"
      End
      Begin VB.Menu hf 
         Caption         =   "数据恢复"
      End
   End
   Begin VB.Menu import_file 
      Caption         =   "帮助"
      Begin VB.Menu line13 
         Caption         =   "-"
      End
      Begin VB.Menu help 
         Caption         =   "操作帮助"
      End
      Begin VB.Menu SS 
         Caption         =   "-"
      End
      Begin VB.Menu TB 
         Caption         =   "填报说明"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'设置主操作窗口最大化显示在mdi窗体
Private Sub maxview_operate_form()
'operate_form.BorderStyle = none
'operate_form.Width = Me.ScaleWidth
'operate_form.Height = Me.ScaleHeight
End Sub
'设置窗体显示在mdi窗体的中央
Sub ShowForm(oForm As Form)
       ' Dim iLeft     As Integer, iTop       As Integer
       ' iLeft = (Me.ScaleWidth - oForm.Width) / 2:               If iLeft < 0 Then iLeft = 0
       ' iTop = (Me.ScaleHeight - oForm.Height) / 2:               If iTop < 0 Then iTop = 0
        'oForm.Left = iLeft:       oForm.Top = iTop
        'oForm.Show
 End Sub
 
Private Sub baobiaoDel_Click()
MainForm.StatusBar1.Panels(2).Text = "状态： 管理历史数据"
Del_history_form.Show
End Sub

Private Sub bf_Click()
Dim f1 As String
Dim f2 As String
f1 = App.Path & "\" & "financialForm.mdb"
If Dir(App.Path & "\" & "备份", vbDirectory) <> "" Then
Else
MkDir App.Path & "\" & "备份"
End If
f2 = App.Path & "\" & "备份" & "\" & "financialForm.mdb"
cn.Close

FileCopy f1, f2
MsgBox "成功备份"
cn.Open
End Sub

Private Sub change_user_Click()
                    Dim choose As Integer
                    choose = MsgBox("你确定要退出,更换其他账号登陆吗?", vbOKCancel)
                    If choose = 1 Then
                     Unload Me
                     login_form.Show
                    Else
                       ' Exit Sub
                    End If


End Sub

Private Sub check_form_Click()
Form1.Show
End Sub



Private Sub fg333_Click()

End Sub

Private Sub help_Click()
App.HelpFile = App.Path & "\help.chm"
SendKeys "{f1}"


End Sub

'Private Sub export_form_Click()
'Load main_form
 '   If userType = "1" Then

  '      Form_Export.nsrbm_value.Caption = main_form.Combo_Nsrbm.Text
       
   ' ElseIf userType = "0" Then
    '    Export_Many_Form.Show
   ' End If
'End Sub

'Private Sub exportExcel_Click()
'Form_Excel.Show
'End Sub

Private Sub import_baobiao_Click()
MainForm.StatusBar1.Panels(2).Text = "状态： 导入报表"
import_baobiao_form.Show
End Sub

Private Sub import_form_info_Click()
MainForm.StatusBar1.Panels(2).Text = "状态： 导入报表数据"
importinit_form.Show

End Sub

Private Sub import_nsrxx_Click()
MainForm.StatusBar1.Panels(2).Text = "状态： 导入纳税人信息"
Unload impnsrxx_form
impnsrxx_form.Show
End Sub

Private Sub info_edit_Click()
MainForm.StatusBar1.Panels(2).Text = "状态： 用户信息维护"
userinfo_form.Show
userinfo_form.text_user_name = username
End Sub

Private Sub line121_Click()

End Sub

Private Sub hf_Click()
impnsrxx_form.Show
End Sub

Private Sub history_Click()
historyBaobiao_form.Show
historyBaobiao_form.date_list.Clear
historyBaobiao_form.nsrbm_value = main_form.Combo_Nsrbm.Text
historyBaobiao_form.nsrmc_valeu = main_form.Label_Nsrmc_Value.Caption

loadDateList
End Sub

Private Sub MDIForm_Load()
MainForm.StatusBar1.Panels(1).Text = "当期操作人员：" & username


End Sub

'修改密码
Private Sub modify_psw_Click()

MainForm.StatusBar1.Panels(2).Text = "状态： 修改密码"
change_psw_form.Show
change_psw_form.lab_name.Caption = username
End Sub

Private Sub quit_sys_Click()
End

End Sub

'显示操作窗口
Private Sub showOperateForm_Click()
   ' operate_form.BorderStyle = none
   ' operate_form.Show
   main_form.Show
End Sub

Private Sub TB_Click()
App.HelpFile = App.Path & "\help2.chm"
SendKeys "{f1}"
End Sub

Private Sub TJNSR_Click()
MainForm.StatusBar1.Panels(2).Text = "状态： 添加新纳税人"
imponsrxx_form.Show


'Unload main_form
'main_form.Show
End Sub

Private Sub version_Click()
MainForm.StatusBar1.Panels(2).Text = "状态： 用户模板设置"
Dim activeFormName As String
Load main_form
    If userType = "1" Then
        version_choose_form.Show
        version_choose_form.Label_Nsrbm = main_form.Combo_Nsrbm.Text
        version_choose_form.Label_Nsrmc = main_form.Label_Nsrmc_Value.Caption
    ElseIf userType = "0" Then
        Version_Form.Show
        
    End If


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
sql = "select * from t_baobiao_content where user_name = '" & username & "' and nsrbm = '" & main_form.Combo_Nsrbm.Text & "'"
Set dateRs = cn.Execute(sql)
While Not dateRs.EOF
    itemFlag = True
    If Trim(dateRs("date_year")) <> "" And Trim(dateRs("date_season")) <> "" Then
         betweenDate = dateRs("date_year") & "年" & change_date(dateRs("date_season")) & "月"
         
         itemCount = historyBaobiao_form.date_list.ListCount - 1
         '去除listview中重复元素
         While itemCount >= 0
            'MsgBox historyBaobiao_form.date_list.List(itemCount) & "" & betweenDate
            If historyBaobiao_form.date_list.List(itemCount) = betweenDate Then
                  itemFlag = False
            End If
            itemCount = itemCount - 1
         Wend
         If itemFlag Then
            historyBaobiao_form.date_list.AddItem betweenDate
         End If
         
    
    End If
    dateRs.MoveNext
Wend
End Sub
