VERSION 5.00
Object = "{B0475000-7740-11D1-BDC3-0020AF9F8E6E}#6.1#0"; "TTF16.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form importinit_form 
   Caption         =   "导入报表数据"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   7500
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   3105
      Left            =   0
      ScaleHeight     =   3045
      ScaleWidth      =   7425
      TabIndex        =   1
      Top             =   0
      Width           =   7485
      Begin VB.CommandButton b_import_init 
         Caption         =   "导入"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3720
         TabIndex        =   5
         Top             =   2160
         Width           =   1695
      End
      Begin VB.CommandButton choose_file 
         Caption         =   "选择..."
         Height          =   495
         Left            =   5520
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox text_validate_num 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   1680
         TabIndex        =   3
         Top             =   1200
         Width           =   3675
      End
      Begin VB.TextBox text_file_path 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   420
         IMEMode         =   3  'DISABLE
         Left            =   1680
         TabIndex        =   2
         Top             =   480
         Width           =   3675
      End
      Begin MSComDlg.CommonDialog init_dia 
         Left            =   720
         Top             =   2040
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "文件路径："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1275
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "激 活 码："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   1290
      End
   End
   Begin TTF160Ctl.F1Book fomular1 
      Height          =   2415
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   4260
      _0              =   $"importinit_form.frx":0000
      _1              =   $"importinit_form.frx":040A
      _2              =   $"importinit_form.frx":0813
      _3              =   $"importinit_form.frx":0C1C
      _4              =   $"importinit_form.frx":1026
      _count          =   5
      _ver            =   2
   End
   Begin MSComDlg.CommonDialog dia 
      Left            =   2160
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "importinit_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


'Private Sub b_import_init_Click()
'
''Dim oXLApplication As New Excel.Application
''Dim oXLWorkbook As Excel.Workbook
''
''Set oXLWorkbook = oXLApplication.Workbooks.Open(App.Path & "\资产负债表.xls")
''
''With oXLWorkbook.Worksheets(2)
''
''MsgBox .Cells(1, 1).value
''
''End With
''oXLWorkbook.Close
''oXLApplication.Quit
'
''getUrl (App.Path & "\现金流量表.xls")
''Dim aaaa As String
''aaaa = getData4IndexSheet(2, 1, 1, 3, 1)
''MsgBox aaaa
''Exit Sub
'
'
'
'
'
'Dim file_path As String  '要导入的文件路径
'Dim validate_num As String  '用户输入的激活码
'Dim validate_num_infile As String '文件中的激活码
'Dim initArray '报表初始数据数组
'
'Dim nsrbmRs As ADODB.Recordset  '用户的数据库的结果集
'Dim nsrxxRs As ADODB.Recordset  '用户的数据库的结果集
'Dim sql As String
'
'Dim success As Integer
'
'file_path = text_file_path.Text
'If file_path <> "" Then
'
'    validate_num = text_validate_num.Text
'    validate_num_infile = getLine(file_path, 1)
'
'    'MsgBox validate_num_infile + "  " + validate_num + "   " + validate_num_infile
'    '验证导入文件的有效性
'    If validate_num = validate_num_infile Then
'      ' MsgBox validate_num + "   " + validate_num_infile
'
'        initArray = getLineArray(file_path)
'        For i = LBound(initArray) + 1 To UBound(initArray)   '循环每一行  即每个报表的信息
'             Dim line As String
'             Dim baobiaoInfoArray  '每行的前四个字段为报表信息
'             Dim nsrbm As String: nsrbm = ""  '纳税人编码
'             Dim baobiaoName As String: baobiaoName = ""   '报表名称
'             Dim version As String: version = ""      '版本
'             Dim betweenDate As String: betweenDate = ""   '所属期
'
'             Dim baobiaoContentID As String    '报表内容ID   根据它关联t_baobiao_value表中的报表单元格值信息
'
'             Dim valuesStr As String
'
'
'             line = initArray(i)
'             If Trim(line) <> "" Then
'
'                 baobiaoInfoArray = Split(line, ",")
'                 nsrbm = baobiaoInfoArray(0)
'                 baobiaoName = baobiaoInfoArray(1)
'                 version = baobiaoInfoArray(2)
'                 betweenDate = baobiaoInfoArray(3)
'
'                 Dim length   '用length表示数组的长度
'                 length = UBound(baobiaoInfoArray) - LBound(baobiaoInfoArray) + 1
'                  MsgBox length
'
'                 ' Exit Sub
'
'                 MsgBox nsrbm & "   " & baobiaoName & "   " & version & "   " & betweenDate
'                 'Exit Sub
'
'
'                 MsgBox App.Path & "\" & baobiaoName & ".xls"
'                 getUrl (App.Path & "\" & baobiaoName & ".xls")
'                ' exportbj As String, nsrbm As String, baobiaoName As String, version As String, betweenTime As String
'                 baobiaoContentID = operate_form.saveBaobiao("1", nsrbm, baobiaoName, version, betweenDate)
'
'                 MsgBox baobiaoContentID
'
'                 If baobiaoContentID = "0" Then
'                    MsgBox "init新建"
'                    baobiaoContentID = operate_form.saveBaobiao("0", nsrbm, baobiaoName, version, betweenDate)  '往t_baobiao_content插入新纪录  返回ID
'                    '保存数据
'                    saveHistoryBb (baobiaoContentID)
'                    savebbValues baobiaoContentID, baobiaoInfoArray
'                 Else
'                    Dim tempFlag As Integer
'                    tempFlag = MsgBox(baobiaoName & "的本期本版本数据已存在，是否覆盖？", 52, "提示")
'                    If tempFlag = 6 Then
'                        '覆盖  保存数据
'                         saveHistoryBb (baobiaoContentID)
'                          savebbValues baobiaoContentID, baobiaoInfoArray
'                    End If
'                 End If
'             End If
'
'        Next
'
'             'Exit Sub
'
'
'    '         If UBound(baobiaoInfoArray) > 0 Then
'    '         nsrbm = baobiaoInfoArray(0)
'    '         nsrqc = baobiaoInfoArray(1)
'    '         End If
'    '
'    '        ' MsgBox nsrbm + "  " + nsrbm
'    '         '数据库操作
'    '        Call check_condatabase
'    '        sql = "select nsrbm from t_nsrxx where nsrbm = '" & nsrbm & "' and username='" & userName & "'"
'    '        Set nsrbmRs = cn.Execute(sql)
'    '        If Not nsrbmRs.EOF Then
'    '            MsgBox "此纳税人信息已经导入！"
'    '            Exit Sub
'    '        End If
'    '        nsrbmRs.Close
'    '        Set nsrbmRs = Nothing
'    '
'    '        sql = "select * from t_nsrxx"
'    '        Set nsrxxRs = New ADODB.Recordset
'    '        Set nsrxxRs.ActiveConnection = cn
'    '        nsrxxRs.LockType = adLockOptimistic
'    '        nsrxxRs.CursorType = adOpenKeyset
'    '        nsrxxRs.Open sql
'    '
'    '        nsrxxRs.AddNew '添加纳税人信息
'    '        nsrxxRs("nsrbm") = nsrbm
'    '        nsrxxRs("nsrqc") = nsrqc
'    '        nsrxxRs("username") = userName
'    '        nsrxxRs.Update
'    '        nsrxxRs.Close
'    '        Set nsrxxRs = Nothing
'    '
'    '       ' success = MsgBox("导入成功！", 1, "提示")
'    '
'    '    Next
'    '    MsgBox ("导入成功！")
'    '    Unload Me
'    '
'    '    operate_form.loadNsrCombox
'    Else
'    MsgBox "激活码不一致！"
'    End If
'Else
'MsgBox "请选择文件！"
'End If
'
'End Sub
'
'Private Sub choose_file_Click()
'On Error GoTo errpro
'Me.dia.InitDir = App.Path
'
'Me.dia.Filter = "文档文件(*.txt)|*.txt"
'
'
'Me.dia.ShowOpen
'
'fileName = Me.dia.fileName
'If fileName = "" Then
'   GoTo errpro
'Else
'   Me.text_file_path.Text = fileName
'   Exit Sub
'End If
'errpro:
''MsgBox "你没有选择任何文件、文件不存在或文件已作废。", vbCritical, "选择错误"
'End Sub
'
''取得excel模版
'Public Function getUrl(theUrl As String)
'If Dir(theUrl) = "" Then '文件不存在
'    MsgBox "报表不存在！"
'Else
'   fomular1.URL = theUrl
'End If
'
'End Function
'
'
'
''取得表上的数据和其坐标，并作为string返回,此方从指定的sheet页中取得数据，适合于多sheet页的excel
''sheetIndex为要取得数据的sheet的序号，从左边为1开始计算
''exlQsdX,exlQsdY为数据的起始点坐标
''exlZdX,exlZdY为数据的终点坐标
'Public Function getData4IndexSheet(sheetIndex As Integer, exlQsdX As Integer, exlQsdY As Integer, exlZdX As Integer, exlZdY As Integer) As String
'
'    Dim copyArray() As Variant
'
'    ReDim copyArray(exlQsdX To exlZdX, exlQsdY To exlZdY) As Variant
'
'    fomular1.CopyDataToArray sheetIndex, exlQsdX, exlQsdY, exlZdX, exlZdY, True, copyArray
'
'   'Action = MsgBox(copyArray(9, 3), vbOKCancel, "ok")
'
'   Dim returnString As String
'   returnString = ""
'
'   For X = exlQsdX To exlZdX
'
'        For Y = exlQsdY To exlZdY
'
'          '  If Trim(copyArray(X, Y)) <> "" Then
'
'                 returnString = returnString & X & "," & Y & "," & Trim(copyArray(X, Y)) & ";"
'
'          '  End If
'
'        Next Y
'
'   Next X
'
'   getData4IndexSheet = returnString
'
'End Function
'
'
''取得表上的数据和其坐标，并作为string返回,此方法默认从第一个sheet页中取得数据，适合于单sheet的excel
''exlQsdX,exlQsdY为数据的起始点坐标
''exlZdX,exlZdY为数据的终点坐标
'Public Function getData(exlQsdX As Integer, exlQsdY As Integer, exlZdX As Integer, exlZdY As Integer) As String
'
'    getData = getData4IndexSheet(1, exlQsdX, exlQsdY, exlZdX, exlZdY)
'
'End Function
'
'Public Function showData(showDataStr As String)
'showData4IndexSheet 1, showDataStr
'End Function
'
''展现数据,此方法可以为指定序号的sheet页展现数据，适合于多sheet页的excel
''sheetIndex为要展现数据的sheet的序号，从左边为1开始计算
'Public Function showData4IndexSheet(sheetIndex As Integer, showDataStr As String)
''isShowData = True
'
''1 第一次解析字符串，array1()数组里面元素为x,y,value （坐标，值）
'Dim array1() As String
'array1() = Split(showDataStr, ";")
'
'
''2 遍历数组处理每个 坐标，值
'Dim array2() As String
'For i = 0 To UBound(array1) - 1
'    '3 第2次解析字符串,对于每一次循环：array2[0]为横坐标，array2[1]为竖坐标，array2[2]为值
'    array2() = Split(array1(i), ",")
'
'
'
'    '-------------------------
'    Dim cellformat As F1CellFormat
'Set cellformat = fomular1.CreateNewCellFormat
'With cellformat
'    .FontColor = vbBlack
'End With
'
'fomular1.Sheet = sheetIndex
'
'fomular1.SetActiveCell array2(0), array2(1)
'
'fomular1.SetCellFormat cellformat
'
'    '-------------------------
'
'    '4 为表格符值
'    fomular1.EntrySRC(sheetIndex, array2(0), array2(1)) = array2(2)
'
'Next i
'
''isShowData = False
'End Function
'
'
'Public Sub saveHistoryBb(bb_content_id As String)
''MsgBox bb_content_id
'operate_form.deleteValue (bb_content_id)  '每次保存都清空记录  全部重新插入
''Exit Sub
'
'Dim dataRange As String   '保存标示输入项的范围的区域  如1，1，3，3 则表明在A1到C3的区域的每个单元格存放的都是一个输入范围。暂存每个报表的sheet2中
'Dim valueRange As String   '保存报表可输入项的范围
'dataRange = getData4IndexSheet(2, 2, 1, 2, 1)   '固定第二行第一列保存可输入项的所有坐标信息
'dataRange = Mid(dataRange, 5, Len(dataRange) - 5)   '  如dataRange = 2,1,12,1,18,1  代表二行一列的值为12，1，18，1  所以这个截串就是获得12,1,18,1
'
''MsgBox dataRange
'If Trim(dataRange) <> "" Then
'
'    Dim dataParamArray
'    dataParamArray = Split(dataRange, ",")
'    'sheet2中保存范围区域参数的数组
'    dataPathArray = Split(dataRange, ",")
'    Dim param1 As Integer
'    Dim param2 As Integer
'    Dim param3 As Integer
'    Dim param4 As Integer
'    Dim param5 As Integer
'    param1 = 2   'sheet2中
'    param2 = CInt(dataParamArray(0))
'    param3 = CInt(dataParamArray(1))
'    param4 = CInt(dataParamArray(2))
'    param5 = CInt(dataParamArray(3))
'    valueRange = getData4IndexSheet(param1, param2, param3, param4, param5)
'    'MsgBox valueRange
'    Dim valuePathArray  '存放sheet1中每段可输入单元格的范围
'    Dim valuePath  '单元格坐标
'    Dim valueStr  As String  'sheet1的可输入框的坐标和值字符串  即getData的返回值  如1,1,ahha;1,2,sdf;
'    valuePathArray = Split(valueRange, ";")
'    For i = LBound(valuePathArray) To UBound(valuePathArray) - 1
'        valuePath = valuePathArray(i)
'        valuePath = Mid(valuePath, 6, Len(valuePath) - 5)  'valuePath = 12,1,12,1,18,1  代表12行1列的值为12，1，18，1  所以这个截串就是获得12,1,18,1
'
'        'MsgBox valuePath
'        If Trim(valuePath) <> "" Then
'
'           'MsgBox "sheet1单元格地址：" & valuePath   'valuePath从数据库中取
'            valueArray = Split(valuePath, ",")
'            param1 = CInt(valueArray(0))
'            param2 = CInt(valueArray(1))
'            param3 = CInt(valueArray(2))
'            param4 = CInt(valueArray(3))
'
'            valueStr = getData(param1, param2, param3, param4)
'            'MsgBox valueStr
'
'           Dim allValueArray  '值数组  如  1,2,asdf;2,3,sdfad;1,3,sdfsf;
'           Dim cellArray   '单元格数组  如  1,2,asdf
'           Dim cell As String
'           Dim row As Integer
'           Dim col As Integer
'           Dim value As String
'
'           allValueArray = Split(valueStr, ";")
'           For j = LBound(allValueArray) To UBound(allValueArray) - 1
'              cell = allValueArray(j)
'              cellArray = Split(cell, ",")
'              row = CInt(cellArray(0))
'              col = CInt(cellArray(1))
'              value = cellArray(2)
'
'              'MsgBox row & "   " & col & "   " & value
'              operate_form.saveValue row, col, value, bb_content_id
'           Next j
'
'        End If
'    Next i
'    MsgBox "保存成功！"
'Else
'MsgBox "此报表没有输入项范围信息，无法保存"
'End If
'
'End Sub
'
'Public Sub savebbValues(bb_content_id As String, baobiaoInfoArray As Variant)   '数组参数为数组  用Variant定义
'MsgBox "进入更新数据" & UBound(baobiaoInfoArray)
'Dim rs As ADODB.Recordset
'Dim strsql As String
'Dim i As Integer   '累加记数
'i = 3
'Dim t As VbMsgBoxResult
'
'Call check_condatabase
'Set rs = New ADODB.Recordset
'Set rs.ActiveConnection = cn
'strsql = "select row_num,col_num,value from t_baobiao_value where bb_content_id = '" & bb_content_id & "' order by col_num,row_num"
'rs.Open strsql, cn, adOpenKeyset, adLockOptimistic
''If Not rs.EOF Then
'While Not rs.EOF
'        i = i + 1
'       ' MsgBox i & "  " & CStr(rs.Fields("row_num")) & "行  " & rs.Fields("col_num") & "列"
'       rs.Fields("value") = baobiaoInfoArray(i)
'       rs.Update
'       rs.MoveNext
'
'Wend
'
'MsgBox i
'
'End Sub
'
Private Sub Picture1_Click()

End Sub
