VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MainForm 
   BackColor       =   &H8000000C&
   Caption         =   "����ָ��ɼ�ϵͳ"
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
      Caption         =   "�û�����"
      Begin VB.Menu line1 
         Caption         =   "-"
      End
      Begin VB.Menu modify_psw 
         Caption         =   "�޸�����"
      End
      Begin VB.Menu line2 
         Caption         =   "-"
      End
      Begin VB.Menu change_user 
         Caption         =   "�л��û�"
      End
      Begin VB.Menu line5 
         Caption         =   "-"
      End
      Begin VB.Menu quit_sys 
         Caption         =   "�˳�ϵͳ"
      End
   End
   Begin VB.Menu option 
      Caption         =   "ģ�����"
      Begin VB.Menu fg33 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu baobiaoDel 
         Caption         =   "������ʷ����"
      End
      Begin VB.Menu fenge87 
         Caption         =   "-"
      End
      Begin VB.Menu history 
         Caption         =   "�鿴��ʷ����"
      End
      Begin VB.Menu feng111 
         Caption         =   "-"
      End
      Begin VB.Menu version 
         Caption         =   "����汾����"
      End
      Begin VB.Menu line88 
         Caption         =   "-"
      End
      Begin VB.Menu showOperateForm 
         Caption         =   "��ʾ����������"
      End
      Begin VB.Menu line12 
         Caption         =   "-"
      End
      Begin VB.Menu exp 
         Caption         =   "��֤����"
         Enabled         =   0   'False
         Visible         =   0   'False
      End
      Begin VB.Menu TJNSR 
         Caption         =   "�����˰��"
      End
   End
   Begin VB.Menu bf_hf 
      Caption         =   "���ݱ�����ָ�"
      Begin VB.Menu line101 
         Caption         =   "-"
      End
      Begin VB.Menu bf 
         Caption         =   "���ݱ���"
      End
      Begin VB.Menu line1211 
         Caption         =   "-"
      End
      Begin VB.Menu hf 
         Caption         =   "���ݻָ�"
      End
   End
   Begin VB.Menu import_file 
      Caption         =   "����"
      Begin VB.Menu line13 
         Caption         =   "-"
      End
      Begin VB.Menu help 
         Caption         =   "��������"
      End
      Begin VB.Menu SS 
         Caption         =   "-"
      End
      Begin VB.Menu TB 
         Caption         =   "�˵��"
      End
   End
End
Attribute VB_Name = "MainForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

'�������������������ʾ��mdi����
Private Sub maxview_operate_form()
'operate_form.BorderStyle = none
'operate_form.Width = Me.ScaleWidth
'operate_form.Height = Me.ScaleHeight
End Sub
'���ô�����ʾ��mdi���������
Sub ShowForm(oForm As Form)
       ' Dim iLeft     As Integer, iTop       As Integer
       ' iLeft = (Me.ScaleWidth - oForm.Width) / 2:               If iLeft < 0 Then iLeft = 0
       ' iTop = (Me.ScaleHeight - oForm.Height) / 2:               If iTop < 0 Then iTop = 0
        'oForm.Left = iLeft:       oForm.Top = iTop
        'oForm.Show
 End Sub
 
Private Sub baobiaoDel_Click()
MainForm.StatusBar1.Panels(2).Text = "״̬�� ������ʷ����"
Del_history_form.Show
End Sub

Private Sub bf_Click()
Dim f1 As String
Dim f2 As String
f1 = App.Path & "\" & "financialForm.mdb"
If Dir(App.Path & "\" & "����", vbDirectory) <> "" Then
Else
MkDir App.Path & "\" & "����"
End If
f2 = App.Path & "\" & "����" & "\" & "financialForm.mdb"
cn.Close

FileCopy f1, f2
MsgBox "�ɹ�����"
cn.Open
End Sub

Private Sub change_user_Click()
                    Dim choose As Integer
                    choose = MsgBox("��ȷ��Ҫ�˳�,���������˺ŵ�½��?", vbOKCancel)
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
MainForm.StatusBar1.Panels(2).Text = "״̬�� ���뱨��"
import_baobiao_form.Show
End Sub

Private Sub import_form_info_Click()
MainForm.StatusBar1.Panels(2).Text = "״̬�� ���뱨������"
importinit_form.Show

End Sub

Private Sub import_nsrxx_Click()
MainForm.StatusBar1.Panels(2).Text = "״̬�� ������˰����Ϣ"
Unload impnsrxx_form
impnsrxx_form.Show
End Sub

Private Sub info_edit_Click()
MainForm.StatusBar1.Panels(2).Text = "״̬�� �û���Ϣά��"
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
MainForm.StatusBar1.Panels(1).Text = "���ڲ�����Ա��" & username


End Sub

'�޸�����
Private Sub modify_psw_Click()

MainForm.StatusBar1.Panels(2).Text = "״̬�� �޸�����"
change_psw_form.Show
change_psw_form.lab_name.Caption = username
End Sub

Private Sub quit_sys_Click()
End

End Sub

'��ʾ��������
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
MainForm.StatusBar1.Panels(2).Text = "״̬�� �������˰��"
imponsrxx_form.Show


'Unload main_form
'main_form.Show
End Sub

Private Sub version_Click()
MainForm.StatusBar1.Panels(2).Text = "״̬�� �û�ģ������"
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

Dim dateRs As ADODB.Recordset  '���汨�������ڵĽ����
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
         betweenDate = dateRs("date_year") & "��" & change_date(dateRs("date_season")) & "��"
         
         itemCount = historyBaobiao_form.date_list.ListCount - 1
         'ȥ��listview���ظ�Ԫ��
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
