VERSION 5.00
Object = "{B0475000-7740-11D1-BDC3-0020AF9F8E6E}#6.1#0"; "TTF16.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form importinit_form 
   Caption         =   "���뱨������"
   ClientHeight    =   3060
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7500
   LinkTopic       =   "Form1"
   ScaleHeight     =   3060
   ScaleWidth      =   7500
   StartUpPosition =   3  '����ȱʡ
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
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "ѡ��..."
         Height          =   495
         Left            =   5520
         TabIndex        =   4
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox text_validate_num 
         BackColor       =   &H00C0FFFF&
         BeginProperty Font 
            Name            =   "����"
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
            Name            =   "����"
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
         Caption         =   "�ļ�·����"
         BeginProperty Font 
            Name            =   "����"
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
         Caption         =   "�� �� �룺"
         BeginProperty Font 
            Name            =   "����"
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
''Set oXLWorkbook = oXLApplication.Workbooks.Open(App.Path & "\�ʲ���ծ��.xls")
''
''With oXLWorkbook.Worksheets(2)
''
''MsgBox .Cells(1, 1).value
''
''End With
''oXLWorkbook.Close
''oXLApplication.Quit
'
''getUrl (App.Path & "\�ֽ�������.xls")
''Dim aaaa As String
''aaaa = getData4IndexSheet(2, 1, 1, 3, 1)
''MsgBox aaaa
''Exit Sub
'
'
'
'
'
'Dim file_path As String  'Ҫ������ļ�·��
'Dim validate_num As String  '�û�����ļ�����
'Dim validate_num_infile As String '�ļ��еļ�����
'Dim initArray '�����ʼ��������
'
'Dim nsrbmRs As ADODB.Recordset  '�û������ݿ�Ľ����
'Dim nsrxxRs As ADODB.Recordset  '�û������ݿ�Ľ����
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
'    '��֤�����ļ�����Ч��
'    If validate_num = validate_num_infile Then
'      ' MsgBox validate_num + "   " + validate_num_infile
'
'        initArray = getLineArray(file_path)
'        For i = LBound(initArray) + 1 To UBound(initArray)   'ѭ��ÿһ��  ��ÿ���������Ϣ
'             Dim line As String
'             Dim baobiaoInfoArray  'ÿ�е�ǰ�ĸ��ֶ�Ϊ������Ϣ
'             Dim nsrbm As String: nsrbm = ""  '��˰�˱���
'             Dim baobiaoName As String: baobiaoName = ""   '��������
'             Dim version As String: version = ""      '�汾
'             Dim betweenDate As String: betweenDate = ""   '������
'
'             Dim baobiaoContentID As String    '��������ID   ����������t_baobiao_value���еı���Ԫ��ֵ��Ϣ
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
'                 Dim length   '��length��ʾ����ĳ���
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
'                    MsgBox "init�½�"
'                    baobiaoContentID = operate_form.saveBaobiao("0", nsrbm, baobiaoName, version, betweenDate)  '��t_baobiao_content�����¼�¼  ����ID
'                    '��������
'                    saveHistoryBb (baobiaoContentID)
'                    savebbValues baobiaoContentID, baobiaoInfoArray
'                 Else
'                    Dim tempFlag As Integer
'                    tempFlag = MsgBox(baobiaoName & "�ı��ڱ��汾�����Ѵ��ڣ��Ƿ񸲸ǣ�", 52, "��ʾ")
'                    If tempFlag = 6 Then
'                        '����  ��������
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
'    '         '���ݿ����
'    '        Call check_condatabase
'    '        sql = "select nsrbm from t_nsrxx where nsrbm = '" & nsrbm & "' and username='" & userName & "'"
'    '        Set nsrbmRs = cn.Execute(sql)
'    '        If Not nsrbmRs.EOF Then
'    '            MsgBox "����˰����Ϣ�Ѿ����룡"
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
'    '        nsrxxRs.AddNew '�����˰����Ϣ
'    '        nsrxxRs("nsrbm") = nsrbm
'    '        nsrxxRs("nsrqc") = nsrqc
'    '        nsrxxRs("username") = userName
'    '        nsrxxRs.Update
'    '        nsrxxRs.Close
'    '        Set nsrxxRs = Nothing
'    '
'    '       ' success = MsgBox("����ɹ���", 1, "��ʾ")
'    '
'    '    Next
'    '    MsgBox ("����ɹ���")
'    '    Unload Me
'    '
'    '    operate_form.loadNsrCombox
'    Else
'    MsgBox "�����벻һ�£�"
'    End If
'Else
'MsgBox "��ѡ���ļ���"
'End If
'
'End Sub
'
'Private Sub choose_file_Click()
'On Error GoTo errpro
'Me.dia.InitDir = App.Path
'
'Me.dia.Filter = "�ĵ��ļ�(*.txt)|*.txt"
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
''MsgBox "��û��ѡ���κ��ļ����ļ������ڻ��ļ������ϡ�", vbCritical, "ѡ�����"
'End Sub
'
''ȡ��excelģ��
'Public Function getUrl(theUrl As String)
'If Dir(theUrl) = "" Then '�ļ�������
'    MsgBox "�������ڣ�"
'Else
'   fomular1.URL = theUrl
'End If
'
'End Function
'
'
'
''ȡ�ñ��ϵ����ݺ������꣬����Ϊstring����,�˷���ָ����sheetҳ��ȡ�����ݣ��ʺ��ڶ�sheetҳ��excel
''sheetIndexΪҪȡ�����ݵ�sheet����ţ������Ϊ1��ʼ����
''exlQsdX,exlQsdYΪ���ݵ���ʼ������
''exlZdX,exlZdYΪ���ݵ��յ�����
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
''ȡ�ñ��ϵ����ݺ������꣬����Ϊstring����,�˷���Ĭ�ϴӵ�һ��sheetҳ��ȡ�����ݣ��ʺ��ڵ�sheet��excel
''exlQsdX,exlQsdYΪ���ݵ���ʼ������
''exlZdX,exlZdYΪ���ݵ��յ�����
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
''չ������,�˷�������Ϊָ����ŵ�sheetҳչ�����ݣ��ʺ��ڶ�sheetҳ��excel
''sheetIndexΪҪչ�����ݵ�sheet����ţ������Ϊ1��ʼ����
'Public Function showData4IndexSheet(sheetIndex As Integer, showDataStr As String)
''isShowData = True
'
''1 ��һ�ν����ַ�����array1()��������Ԫ��Ϊx,y,value �����ֵ꣬��
'Dim array1() As String
'array1() = Split(showDataStr, ";")
'
'
''2 �������鴦��ÿ�� ���ֵ꣬
'Dim array2() As String
'For i = 0 To UBound(array1) - 1
'    '3 ��2�ν����ַ���,����ÿһ��ѭ����array2[0]Ϊ�����꣬array2[1]Ϊ�����꣬array2[2]Ϊֵ
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
'    '4 Ϊ����ֵ
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
'operate_form.deleteValue (bb_content_id)  'ÿ�α��涼��ռ�¼  ȫ�����²���
''Exit Sub
'
'Dim dataRange As String   '�����ʾ������ķ�Χ������  ��1��1��3��3 �������A1��C3�������ÿ����Ԫ���ŵĶ���һ�����뷶Χ���ݴ�ÿ�������sheet2��
'Dim valueRange As String   '���汨���������ķ�Χ
'dataRange = getData4IndexSheet(2, 2, 1, 2, 1)   '�̶��ڶ��е�һ�б���������������������Ϣ
'dataRange = Mid(dataRange, 5, Len(dataRange) - 5)   '  ��dataRange = 2,1,12,1,18,1  �������һ�е�ֵΪ12��1��18��1  ��������ش����ǻ��12,1,18,1
'
''MsgBox dataRange
'If Trim(dataRange) <> "" Then
'
'    Dim dataParamArray
'    dataParamArray = Split(dataRange, ",")
'    'sheet2�б��淶Χ�������������
'    dataPathArray = Split(dataRange, ",")
'    Dim param1 As Integer
'    Dim param2 As Integer
'    Dim param3 As Integer
'    Dim param4 As Integer
'    Dim param5 As Integer
'    param1 = 2   'sheet2��
'    param2 = CInt(dataParamArray(0))
'    param3 = CInt(dataParamArray(1))
'    param4 = CInt(dataParamArray(2))
'    param5 = CInt(dataParamArray(3))
'    valueRange = getData4IndexSheet(param1, param2, param3, param4, param5)
'    'MsgBox valueRange
'    Dim valuePathArray  '���sheet1��ÿ�ο����뵥Ԫ��ķ�Χ
'    Dim valuePath  '��Ԫ������
'    Dim valueStr  As String  'sheet1�Ŀ������������ֵ�ַ���  ��getData�ķ���ֵ  ��1,1,ahha;1,2,sdf;
'    valuePathArray = Split(valueRange, ";")
'    For i = LBound(valuePathArray) To UBound(valuePathArray) - 1
'        valuePath = valuePathArray(i)
'        valuePath = Mid(valuePath, 6, Len(valuePath) - 5)  'valuePath = 12,1,12,1,18,1  ����12��1�е�ֵΪ12��1��18��1  ��������ش����ǻ��12,1,18,1
'
'        'MsgBox valuePath
'        If Trim(valuePath) <> "" Then
'
'           'MsgBox "sheet1��Ԫ���ַ��" & valuePath   'valuePath�����ݿ���ȡ
'            valueArray = Split(valuePath, ",")
'            param1 = CInt(valueArray(0))
'            param2 = CInt(valueArray(1))
'            param3 = CInt(valueArray(2))
'            param4 = CInt(valueArray(3))
'
'            valueStr = getData(param1, param2, param3, param4)
'            'MsgBox valueStr
'
'           Dim allValueArray  'ֵ����  ��  1,2,asdf;2,3,sdfad;1,3,sdfsf;
'           Dim cellArray   '��Ԫ������  ��  1,2,asdf
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
'    MsgBox "����ɹ���"
'Else
'MsgBox "�˱���û�������Χ��Ϣ���޷�����"
'End If
'
'End Sub
'
'Public Sub savebbValues(bb_content_id As String, baobiaoInfoArray As Variant)   '�������Ϊ����  ��Variant����
'MsgBox "�����������" & UBound(baobiaoInfoArray)
'Dim rs As ADODB.Recordset
'Dim strsql As String
'Dim i As Integer   '�ۼӼ���
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
'       ' MsgBox i & "  " & CStr(rs.Fields("row_num")) & "��  " & rs.Fields("col_num") & "��"
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
