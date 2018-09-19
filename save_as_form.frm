VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form save_as_form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "另存为EXCEL"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6840
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   6840
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton b_quit 
      Caption         =   "退出"
      Height          =   495
      Left            =   4920
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Export 
      Caption         =   "保存"
      Height          =   495
      Left            =   3360
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton Cmd_Choose 
      Caption         =   "选择路径.."
      Height          =   375
      Left            =   4800
      TabIndex        =   2
      Top             =   960
      Width           =   1215
   End
   Begin VB.TextBox text_import_path 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1440
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin MSComDlg.CommonDialog CommonDialog_Export 
      Left            =   240
      Top             =   2280
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "另  存  为："
      Height          =   255
      Left            =   240
      TabIndex        =   0
      Top             =   1080
      Width           =   1095
   End
End
Attribute VB_Name = "save_as_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub b_quit_Click()
Unload Me
End Sub

Private Sub Cmd_Choose_Click()
Dim filePath As String

filePath = GetFolder("打开一个目录", Me.CommonDialog_Export)
Me.text_import_path.Text = filePath
End Sub

Private Sub Cmd_Export_Click()
Dim bbbb As String
Dim xy As String
Dim bj1 As String
Dim bj2 As String

   If Me.text_import_path.Text <> "" Then
  
   
   ' main_form.CB_Season.Text = date_change(main_form.CB_Season.Text)
     If Me.text_import_path.Text <> "C:\" And Me.text_import_path.Text <> "D:\" And Me.text_import_path.Text <> "E:\" And Me.text_import_path.Text <> "F:\" And Me.text_import_path.Text <> "G:\" And Me.text_import_path.Text <> "H:\" Then
    '判断是不是根目录。
    
    xy = Me.text_import_path.Text & "\" & main_form.Combo_Nsrbm.Text & "_" & main_form.CB_Year.Text & "_" & main_form.CB_Season.Text
    Else
    xy = Me.text_import_path.Text & main_form.Combo_Nsrbm.Text & "_" & main_form.CB_Year.Text & "_" & main_form.CB_Season.Text
    End If
    
  '  xy = Me.text_import_path.Text & "\" & main_form.Combo_Nsrbm.Text & "_" & main_form.CB_Year.Text & "_" & main_form.CB_Season.Text
     If Dir(xy & ".xls") <> "" Then
  Kill xy & ".xls"
  Else
  End If
    
    
 
   main_form.F1Book1.Sheet = 1
   
   If main_form.F1Book1.ObjValue(pid) = -1 Or main_form.F1Book1.ObjValue(pid2) = -1 Then
   MsgBox "经营状况或者征收机构没有填写"
   Exit Sub
   Else
   main_form.F1Book1.WriteEx xy & ".xls", F1FileExcel97
  ' main_form.F1Book1.Sheet = 1
  ' bj1 = main_form.F1Book1.ObjItem(pid, main_form.F1Book1.ObjValue(pid))
  ' bj2 = main_form.F1Book1.ObjItem(pid2, main_form.F1Book1.ObjValue(pid2))
  '  Set xlBook = xlApp.Workbooks().Open(xy & ".xls")
 '    ActiveSheet.Unprotect
   '    Sheets("经营信息表").Select
   '    Cells(7, 3) = bj1
  '     Cells(8, 3) = bj2
   '   ActiveWorkbook.save
 '  ActiveWorkbook.Close
   'main_form.CB_Season.Text = change_date(main_form.CB_Season.Text)
   MsgBox "另存成功"
   Unload Me
   End If
   
 Else
   MsgBox "请选择导出路径！"
   End If
   
End Sub
