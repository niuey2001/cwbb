VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form Form_Excel 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "导出Excle"
   ClientHeight    =   2730
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6825
   Icon            =   "Form_Excel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2730
   ScaleWidth      =   6825
   Begin VB.CommandButton Cmd_Export 
      Caption         =   "导出"
      Height          =   495
      Left            =   3120
      TabIndex        =   4
      Top             =   2040
      Width           =   1215
   End
   Begin VB.CommandButton b_quit 
      Caption         =   "退出"
      Height          =   495
      Left            =   4680
      TabIndex        =   3
      Top             =   2040
      Width           =   1215
   End
   Begin VB.TextBox text_import_path 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   960
      Width           =   3255
   End
   Begin VB.CommandButton Cmd_Choose 
      Caption         =   "选择路径.."
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   960
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CommonDialog_Export 
      Left            =   4080
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label2 
      Caption         =   "导  出  至："
      Height          =   255
      Left            =   0
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
End
Attribute VB_Name = "Form_Excel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub b_quit_Click()
'excel.Application.Quit
Unload Me
End Sub

Private Sub Cmd_Choose_Click()
Dim filePath As String

filePath = GetFolder("打开一个目录", Me.CommonDialog_Export)
Me.text_import_path.Text = filePath
End Sub
Public Function jiami(X As String) '加密处理:通过2进制下的异或运算来实现加密功能。异或运算值:18
Dim B() As Byte
Dim PassWord As String
Dim B1 As Byte
Dim i As Long, l As Long, j As Long
Open X For Binary As #1
If LOF(1) > 0 Then
j = LOF(1)
ReDim B(LOF(1) - 1)
Get #1, , B
End If
Close #1

'Dim P As Long
'PassWord = "OutsideFile"
'l = Len(PassWord)

'ReDim B1(l)
'For I = 1 To l
'B1(I) = Asc(Mid(PassWord, I, 1))

'Next
B1 = 18
For i = 0 To UBound(B)
B(i) = B(i) Xor B1

Next

Open X For Binary As #1
Put #1, , B
Close #1

End Function

Private Sub Cmd_Export_Click()
Dim bbbb As String
Dim xy As String
 Dim rarpath As String
   Dim path1 As String
   Dim path2 As String
   Dim path3 As String
   Dim path4 As String
   Dim sql As String


 Dim hy_rs As ADODB.Recordset
 
 Dim v_hy As String
 
 Call check_condatabase
 
 sql = "select zchy from t_nsrxx where  nsrbm='" & main_form.Combo_Nsrbm.Text & " ' and   username='" & username & "'"
 
 
 Set hy_rs = cn.Execute(sql)
 
 v_hy = hy_rs("zchy")
 

   If Me.text_import_path.Text <> "" Then
  
    Dim x_Rs As ADODB.Recordset

    bbbb = main_form.Label_Bb_Value.Caption
    
    
   
    main_form.CB_Season.Text = date_change(main_form.CB_Season.Text)
    
    Call check_condatabase
sql = "select id from t_baobiao_content where nsrbm='" & main_form.Combo_Nsrbm.Text & " 'and date_year='" & main_form.CB_Year.Text & "' and date_Season='" & main_form.CB_Season.Text & "'"
Set x_Rs = cn.Execute(sql)
If x_Rs.EOF Then
MsgBox ("请保存后在导出")
main_form.CB_Season.Text = change_date(main_form.CB_Season.Text)

Else

    If validate_exp_data(main_form.CB_Year.Text, main_form.CB_Season.Text, main_form.Combo_Nsrbm.Text, Form_Export.text_wrong_mes, Form_Export.ScriptControl1, main_form.F1Book1) Then
    If Me.text_import_path.Text <> "C:\" And Me.text_import_path.Text <> "D:\" And Me.text_import_path.Text <> "E:\" And Me.text_import_path.Text <> "F:\" And Me.text_import_path.Text <> "G:\" And Me.text_import_path.Text <> "H:\" Then
    '判断是不是根目录。
    
    xy = Me.text_import_path.Text & "\" & main_form.Combo_Nsrbm.Text & "_" & main_form.CB_Year.Text & "_" & main_form.CB_Season.Text
    Else
    xy = Me.text_import_path.Text & main_form.Combo_Nsrbm.Text & "_" & main_form.CB_Year.Text & "_" & main_form.CB_Season.Text
    End If
    main_form.F1Book1.WriteEx xy & ".xls", F1FileExcel97
    If bbbb = "2007年版" Then
       bbbb = "01"
     ElseIf bbbb = "2005年版" Then
       bbbb = "02"
    End If
  ' Dim xy As String
   'Dim xlApp As New excel.Application
   Set xlBook = xlApp.Workbooks().Open(xy & ".xls") '分开存放4个人表
   Sheets("资产负债表").Select
   Sheets("资产负债表").Copy
   ActiveSheet.Unprotect
   Cells.Select
        Selection.Copy
        
        Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=False '去除表中的公式
  If Dir(xy & "_" & bbbb & "_" & "ZCFZB" & "_" & ".xls") <> "" Then
  Kill xy & "_" & bbbb & "_" & "ZCFZB" & "_" & ".xls"
  Else
  End If
   ActiveWorkbook.SaveAs fileName:=xy & "_" & bbbb & "_" & "ZCFZB" & "_" & ".xls"
   ActiveWorkbook.Close
   jiami (xy & "_" & bbbb & "_" & "ZCFZB" & "_" & ".xls")
   Sheets("经营信息表").Select
   Sheets("经营信息表").Copy
   ActiveSheet.Unprotect
   Cells.Select
        Selection.Copy
        
        Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=False
 If Dir(xy & "_" & bbbb & "_" & "JYXXB" & "_" & ".xls") <> "" Then
    Kill xy & "_" & bbbb & "_" & "JYXXB" & "_" & ".xls"
Else
End If
   ActiveWorkbook.SaveAs fileName:=xy & "_" & bbbb & "_" & "JYXXB" & "_" & ".xls"
   ActiveWorkbook.Close
   jiami (xy & "_" & bbbb & "_" & "JYXXB" & "_" & ".xls")
   Sheets("利润表").Select
   Sheets("利润表").Copy
   ActiveSheet.Unprotect
   Cells.Select
        Selection.Copy
        
        Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=False
 If Dir(xy & "_" & bbbb & "_" & "LRB" & "_" & ".xls") <> "" Then
    Kill xy & "_" & bbbb & "_" & "LRB" & "_" & ".xls"
 Else
 End If
   ActiveWorkbook.SaveAs fileName:=xy & "_" & bbbb & "_" & "LRB" & "_" & ".xls"
   ActiveWorkbook.Close
   jiami (xy & "_" & bbbb & "_" & "LRB" & "_" & ".xls")
   Sheets("现金流量表").Select
   Sheets("现金流量表").Copy
   ActiveSheet.Unprotect
   Cells.Select
        Selection.Copy
        
        Selection.PasteSpecial Paste:=xlValues, Operation:=xlNone, SkipBlanks:= _
                False, Transpose:=False
  If Dir(xy & "_" & bbbb & "_" & "XJLLB" & "_" & ".xls") <> "" Then
   Kill xy & "_" & bbbb & "_" & "XJLLB" & "_" & ".xls"
  Else
  End If
   ActiveWorkbook.SaveAs fileName:=xy & "_" & bbbb & "_" & "XJLLB" & "_" & ".xls"
   ActiveWorkbook.Close
   jiami (xy & "_" & bbbb & "_" & "XJLLB" & "_" & ".xls")
   ActiveWorkbook.save
   ActiveWorkbook.Close
  
   path1 = Chr(34) & xy & "_" & bbbb & "_" & "ZCFZB" & "_" & ".xls" & Chr(34) '通过char(34)解决路径中带空格的问题
  
   path2 = Chr(34) & xy & "_" & bbbb & "_" & "JYXXB" & "_" & ".xls" & Chr(34)
   path3 = Chr(34) & xy & "_" & bbbb & "_" & "LRB" & "_" & ".xls" & Chr(34)
   path4 = Chr(34) & xy & "_" & bbbb & "_" & "XJLLB" & "_" & ".xls" & Chr(34)
   
   rarpath = Chr(34) & xy & "_" & bbbb & "_" & "new" & "_" & v_hy & ".zip" & Chr(34)
   
   Shell App.Path & "\winrar.exe a -ep -afzip " & rarpath & " " & path1 & " " & path2 & " " & path3 & " " & path4

   main_form.CB_Season.Text = change_date(main_form.CB_Season.Text)
   MsgBox "导出成功！"


   Kill xy & "_" & bbbb & "_" & "ZCFZB" & "_" & ".xls"
   Kill xy & "_" & bbbb & "_" & "JYXXB" & "_" & ".xls"
   Kill xy & "_" & bbbb & "_" & "LRB" & "_" & ".xls"
   Kill xy & "_" & bbbb & "_" & "XJLLB" & "_" & ".xls"
   
   Kill xy & ".xls"
   Else
MsgBox "该表格数据验证没通过"
main_form.CB_Season.Text = change_date(main_form.CB_Season.Text)
End If
Unload Me

End If

 Else
  MsgBox "请选择导出路径！"
 End If
 

End Sub
