VERSION 5.00
Object = "{B0475000-7740-11D1-BDC3-0020AF9F8E6E}#6.1#0"; "TTF16.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form operate_form 
   Caption         =   "����ָ��ɼ�����"
   ClientHeight    =   8490
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   566
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   792
   Visible         =   0   'False
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   255
      Left            =   17760
      TabIndex        =   28
      Top             =   9960
      Width           =   615
   End
   Begin VB.CommandButton loadData 
      Caption         =   "�鿴��ʷ����"
      Height          =   495
      Left            =   9000
      TabIndex        =   14
      Top             =   360
      Width           =   1335
   End
   Begin VB.CommandButton export_data 
      Caption         =   "����"
      Height          =   495
      Left            =   5520
      TabIndex        =   13
      Top             =   120
      Width           =   1095
   End
   Begin VB.CommandButton save 
      Caption         =   "����"
      Height          =   495
      Left            =   10320
      TabIndex        =   12
      Top             =   720
      Width           =   1095
   End
   Begin VB.CommandButton loadBaobiaoValue 
      Caption         =   "���س�ʼ����"
      Height          =   495
      Left            =   7200
      TabIndex        =   11
      Top             =   240
      Width           =   1335
   End
   Begin VB.ComboBox CB_EndYear 
      Height          =   300
      Left            =   7800
      TabIndex        =   10
      Top             =   720
      Width           =   735
   End
   Begin VB.ComboBox CB_EndMonth 
      Height          =   300
      Left            =   8760
      TabIndex        =   9
      Top             =   720
      Width           =   615
   End
   Begin VB.ComboBox CB_StartYear 
      Height          =   300
      Left            =   5640
      TabIndex        =   8
      Top             =   720
      Width           =   735
   End
   Begin VB.ComboBox CB_StartMonth 
      Height          =   300
      Left            =   6600
      TabIndex        =   7
      Top             =   720
      Width           =   615
   End
   Begin VB.Frame sgy_info_frame 
      Caption         =   "��˰�˻�����Ϣ"
      Height          =   2655
      Left            =   120
      TabIndex        =   2
      Top             =   360
      Width           =   4095
      Begin VB.TextBox text_nsrqc 
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   1680
         Width           =   3855
      End
      Begin VB.ComboBox combox_nsrbm 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   720
         Width           =   3855
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   1920
         Top             =   2160
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label4 
         Caption         =   "��˰��ȫ�ƣ�"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   1320
         Width           =   1215
      End
      Begin VB.Label Label3 
         Caption         =   "��˰�˱��룺"
         Height          =   375
         Left            =   120
         TabIndex        =   5
         Top             =   360
         Width           =   1455
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "�����б�"
      Height          =   5175
      Left            =   120
      TabIndex        =   0
      Top             =   3000
      Width           =   4095
      Begin VB.ListBox baobiaoList 
         BackColor       =   &H00FFFFFF&
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3630
         Left            =   360
         TabIndex        =   1
         Top             =   360
         Width           =   3855
      End
   End
   Begin TTF160Ctl.F1Book F1Book1 
      Height          =   7215
      Left            =   4440
      TabIndex        =   15
      Top             =   1200
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   12726
      _0              =   $"operate_form.frx":0000
      _1              =   $"operate_form.frx":0409
      _2              =   $"operate_form.frx":0812
      _3              =   $"operate_form.frx":0C1B
      _4              =   $"operate_form.frx":1024
      _count          =   5
      _ver            =   2
   End
   Begin VB.Label Label2 
      Caption         =   "__"
      Height          =   255
      Left            =   7560
      TabIndex        =   27
      Top             =   720
      Width           =   255
   End
   Begin VB.Label Label10 
      Caption         =   "�������ͣ�"
      Height          =   375
      Left            =   9240
      TabIndex        =   26
      Top             =   240
      Width           =   975
   End
   Begin VB.Label lable_version 
      Caption         =   "�°�"
      Height          =   255
      Left            =   10200
      TabIndex        =   25
      Top             =   240
      Width           =   495
   End
   Begin VB.Label Label12 
      Caption         =   "��λ��Ԫ"
      Height          =   255
      Left            =   10920
      TabIndex        =   24
      Top             =   240
      Width           =   855
   End
   Begin VB.Label Label15 
      Caption         =   "¼��ʱ�䣺"
      Height          =   255
      Left            =   12960
      TabIndex        =   23
      Top             =   9840
      Width           =   975
   End
   Begin VB.Label Label13 
      Caption         =   "¼����Ա��"
      Height          =   255
      Left            =   10320
      TabIndex        =   22
      Top             =   9840
      Width           =   975
   End
   Begin VB.Label Label9 
      Caption         =   "��"
      Height          =   255
      Left            =   8520
      TabIndex        =   21
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label8 
      Caption         =   "��"
      Height          =   255
      Left            =   9360
      TabIndex        =   20
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "��"
      Height          =   255
      Left            =   6360
      TabIndex        =   19
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label6 
      Caption         =   "��"
      Height          =   255
      Left            =   7200
      TabIndex        =   18
      Top             =   720
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "���������ڼ䣺"
      Height          =   255
      Left            =   4320
      TabIndex        =   17
      Top             =   720
      Width           =   1335
   End
   Begin VB.Label lable_bb_name 
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   18
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   9360
      TabIndex        =   16
      Top             =   0
      Width           =   2295
   End
End
Attribute VB_Name = "operate_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Public baobiaoEditBj  As String   '�����б�û��ѡ��ʱ
'
'
'
'
'Public Sub viewVersion()
'    Dim version As String
'
'    Dim versionStr As String
'    versionStr = getData4IndexSheet(2, 1, 1, 1, 1)
'    version = Mid(versionStr, 5, 1)
'    If version = "1" Then
'        lable_version.Caption = "�°�"
'    Else
'        lable_version.Caption = "�ɰ�"
'    End If
'
'
'End Sub
'
'Private Sub baobiaoList_DblClick()
'baobiaoEditBj = "1"
'Dim baobiaoName As String
'baobiaoName = baobiaoList.Text
'lable_bb_name.Caption = baobiaoName
'getUrl (App.Path & "\" & baobiaoName & ".xls")
'viewVersion  '��ʾ�汾
'End Sub
'
'Private Sub combox_nsrbm_Click()
'Dim nsrRs As ADODB.Recordset '������˰�˵Ľ����
'Dim sql As String
'
'Dim nsrbmText As String
'Dim nsrqcText As String
'nsrbmText = Trim(combox_nsrbm.Text)
'
'Call check_condatabase
'If nsrbmText <> "" Then
'    sql = "select nsrqc from t_nsrxx where username='" & userName & "' and nsrbm ='" & nsrbmText & "'"
'    Set nsrRs = cn.Execute(sql)
'    If Not nsrRs.EOF Then
'        nsrRs.MoveFirst
'        nsrqcText = nsrRs("nsrqc")
'        Me.text_nsrqc.Text = nsrqcText
'    End If
'End If
'
'End Sub
'
'Private Sub Command2_Click()
'
'End Sub
'
''������������   ����֮ǰҪ��֤��δʵ�֣�
'Private Sub export_data_Click()
'Dim mes As String
'mes = option_validate
'If mes <> "" Then
'    MsgBox mes
'    Exit Sub
'End If
'
'Dim nsrbm As String
'Dim baobiaoName As String
'Dim version As String  '�汾
'Dim betweenTime As String  '���ñ��
'
'nsrbm = combox_nsrbm.Text
'baobiaoName = lable_bb_name.Caption
'version = lable_version.Caption
'If version = "�°�" Then
'    version = "1"
'ElseIf version = "�ɰ�" Then
'    version = "0"
'End If
'betweenTime = CB_StartYear.Text & CB_StartMonth.Text & "-" & CB_EndYear.Text & CB_EndMonth.Text
'
'
'Dim headinfo As String
'Dim bb_content_id As String
'Dim allValueStr As String
'
'bb_content_id = saveBaobiao("1", nsrbm, baobiaoName, version, betweenTime) '���ر�������ID  �������򷵻�"0"
''MsgBox bb_content_id
'If bb_content_id <> "0" Then
'    allValueStr = getExportValuesById(bb_content_id)  '  1,1,sdsd;2,1,asd;1,2,dfad;2,2,dfa;
'    Set fileObj = CreateObject("Scripting.FileSystemObject")
'   ' Set writeObj = fileObj.CreateTextFile(App.Path & "\export\����������.txt", True)
'
'    Set writeObj = fileObj.CreateTextFile("D:\����������.txt", True)
'
'    headinfo = nsrbm & "," & baobiaoName & "," & version & "," & betweenTime
'   ' MsgBox headinfo
'    writeObj.WriteLine (headinfo & "," & allValueStr)
'    writeObj.Close
'    MsgBox "�����ɹ������ݵ�����D�̸�·���£�����������.txt��"
'Else
'    MsgBox "���ڱ���û�б��棡"
'End If
'
'End Sub
'
'Private Sub Form_Load()
'  Me.Width = ScaleX(800, vbPixels, vbTwips)   '�趨����Ŀ��Ϊ800����
'  Me.Height = ScaleY(600, vbPixels, vbTwips)  '�趨����ĸ߶�Ϊ680����
'
''Me.WindowState = 2
'
'
'
'baobiaoEditBj = "0"   'ҳ���ʼ��ʱ  �����б�û�б�ѡ��
'loadNsrCombox
'loadBaobiao
'loadDate  '��������������
'
'End Sub
''��������
'Public Sub loadDate()
'  Dim i As Integer
''  Me.CB_StartYear.Text = CStr(Year(Date))
''  Me.CB_EndYear.Text = CStr(Year(Date))
''  If Month(Date) = 1 Then
''    Me.CB_StartMonth.Text = "12"
''    Me.CB_StartYear.Text = CStr(Year(Date) - 1)
''  Else
''    Me.CB_StartMonth.Text = CStr(Month(Date) - 1)
''  End If
''
''  Me.CB_EndMonth.Text = CStr(Month(Date))
''  Me.CB_StartDay.Text = CStr(Day(Date))
''  Me.CB_EndDay.Text = CStr(Day(Date))
'
'  For i = 2000 To 2100
'    Me.CB_StartYear.AddItem CStr(i)
'    Me.CB_EndYear.AddItem CStr(i)
'  Next i
'  For i = 1 To 12
'    If i < 10 Then
'        Me.CB_StartMonth.AddItem "0" & CStr(i)
'        Me.CB_EndMonth.AddItem "0" & CStr(i)
'    Else
'        Me.CB_StartMonth.AddItem CStr(i)
'        Me.CB_EndMonth.AddItem CStr(i)
'    End If
'
'
'  Next i
'End Sub
''��������ҳ��ʼ�����
'Private Sub Form_Resize()
' 'If Me.WindowState <> 2 Then
' '      Me.WindowState = 2
' 'End If
'
'End Sub
'
'Public Sub loadNsrCombox()
'
''If userType = "1" Then
'    combox_nsrbm.Clear
'    Dim nsrRs As ADODB.Recordset '������˰�˵Ľ����
'    Dim sql As String
'
'    Call check_condatabase
'    sql = "select nsrbm from t_nsrxx where username='" & userName & "'"
'    Set nsrRs = cn.Execute(sql)
'    While Not nsrRs.EOF
'        If Trim(nsrRs("nsrbm")) <> "" Then
'        Me.combox_nsrbm.AddItem nsrRs("nsrbm")
'        End If
'        nsrRs.MoveNext
'    Wend
'
'    If combox_nsrbm.ListCount > 0 Then
'    Me.combox_nsrbm.ListIndex = 0
'    End If
'
''End If
'
'
''combox_nsrbm rs.movefirst
'End Sub
'
'Public Sub loadBaobiao()
'baobiaoList.Clear
'Dim baobiaoRs As ADODB.Recordset  '���汨�����ƵĽ����
'Dim sql As String
'
'Call check_condatabase
'sql = "select baobiao_name from t_baobiao where bj= '1'"
'Set baobiaoRs = cn.Execute(sql)
'While Not baobiaoRs.EOF
'    If Trim(baobiaoRs("baobiao_name")) <> "" Then
'         Me.baobiaoList.AddItem baobiaoRs("baobiao_name")
'    End If
'    baobiaoRs.MoveNext
'Wend
'
'End Sub
''ȡ��excelģ��
'Public Function getUrl(theUrl As String)
'If Dir(theUrl) = "" Then '�ļ�������
'    MsgBox "�������ڣ�"
'Else
'   F1Book1.URL = theUrl
'End If
'
'End Function
'
'Public Function option_validate()
''Dim ok_flag As Boolean
'If Me.combox_nsrbm.Text = "" Then
'   option_validate = "���ȵ�����˰����Ϣ��"
'ElseIf baobiaoEditBj = "0" Then
'   option_validate = "����ѡ����౨���б�"
'ElseIf Trim(CB_StartYear.Text) = "" Or Trim(CB_StartMonth.Text) = "" Or Trim(CB_EndYear.Text) = "" Or Trim(CB_EndMonth.Text) = "" Then
'     option_validate = "���������ڲ���Ϊ�գ�"
'Else
'     option_validate = ""
'End If
'
'
'
'
'
'End Function
'
'Private Sub Label14_Click()
'
'End Sub
'
'Private Sub loadBaobiaoValue_Click()
'Dim mes As String
'mes = option_validate
'If mes <> "" Then
'    MsgBox mes
'    Exit Sub
'End If
'
'Dim baobiaoValueRs As ADODB.Recordset '������˰�˵Ľ����
'Dim sql As String
'Dim nsrbm As String
'Dim baobiaoName As String
'Dim version As String
'Dim between_date As String   '200901-200912   201001-201004
'
'
'
'Dim id As String
'id = "0"
'nsrbm = combox_nsrbm.Text
'baobiaoName = lable_bb_name.Caption
'version = lable_version.Caption
'
'If Trim(version) = "�°�" Then
'    version = "1"
'Else
'    version = "0"
'End If
'
'Dim startYear As String
'Dim startYearInt As Integer
'startYear = CB_StartYear.Text
'
'
'Dim initBj As String  'initBj ������ĳ�ʼ������   ���������  ���� �ڳ��� 0�����޳�ʼ������  1���������   2�����ڳ���
'initBj = getData4IndexSheet(2, 3, 1, 3, 1)
'initBj = Mid(initBj, 5, 1)
'
'
'If initBj = "1" And startYear <> "" Then
'startYearInt = CInt(startYear)
'startYear = CStr(startYearInt - 1)
'between_date = startYear & "01" & "-" & startYear & "12"
''MsgBox between_date
'End If
'
''between_date = CB_StartYear.Text & CB_StartMonth.Text & "-" & CB_EndYear.Text & CB_EndMonth.Text
'
'Call check_condatabase
'sql = "select id from t_baobiao_content where nsrbm = '" & nsrbm & "' and baobiao_name='" & baobiaoName & "' and version = '" & version & "' and bb_between_date = '" & between_date & "'"
''MsgBox sql
'
'Set baobiaoValueRs = cn.Execute(sql)
'
''    If Trim(nsrRs("nsrbm")) <> "" Then
''    Me.combox_nsrbm.AddItem nsrRs("nsrbm")
''    End If
''    nsrRs.MoveNext
''Wend
'If Not baobiaoValueRs.EOF Then
'    id = CStr(baobiaoValueRs("id"))
'    'MsgBox id
'End If
'baobiaoValueRs.Close
'Set baobiaoValueRs = Nothing
'
'If id <> "0" Then
'       Dim allValueStr As String
'       allValueStr = getInitValuesById(id)
'       showData (allValueStr)
'Else
'      MsgBox "û�г�ʼ�����ݣ�"
'End If
'
'
'End Sub
'Public Function getInitValuesById(id As String)
'    Dim valueRs As ADODB.Recordset '���汨������������
'    Dim rowNum As String
'    Dim colNum As String
'    Dim value As String
'    Dim allValueStr As String
'
'    Dim initColNum As String
'    Dim initColNumArray
'    initColNum = getData4IndexSheet(2, 4, 1, 4, 1)
'    initColNum = Mid(initColNum, 5, Len(initColNum) - 5)
'    'MsgBox initColNum
'    initColNumArray = Split(initColNum, ",")
'
'    allValueStr = ""
'    Call check_condatabase
'    sql = "select row_num,col_num,value from t_baobiao_value where bb_content_id = '" & id & "'"
'    Set valueRs = cn.Execute(sql)
'
'    While Not valueRs.EOF
'
'        rowNum = CStr(valueRs("row_num"))
'        colNum = CStr(valueRs("col_num"))
'        value = valueRs("value")
'
'        For i = LBound(initColNumArray) To UBound(initColNumArray)
'            If colNum = initColNumArray(i) Then
'                allValueStr = allValueStr + rowNum & "," & colNum & "," & value & ";"
'            End If
'        Next
'
'
'        valueRs.MoveNext
'    Wend
'    valueRs.Close
'    Set valueRs = Nothing
'    'MsgBox allValueStr
'    getInitValuesById = allValueStr
'End Function
'Public Function getValuesById(id As String)
'    Dim valueRs As ADODB.Recordset '���汨������������
'    Dim rowNum As String
'    Dim colName As String
'    Dim value As String
'
'    allValueStr = ""
'    Call check_condatabase
'    sql = "select row_num,col_num,value from t_baobiao_value where bb_content_id = '" & id & "'"
'    Set valueRs = cn.Execute(sql)
'
'    While Not valueRs.EOF
'        rowNum = CStr(valueRs("row_num"))
'        colNum = CStr(valueRs("col_num"))
'        value = valueRs("value")
'
'        allValueStr = allValueStr + rowNum & "," & colNum & "," & value & ";"
'        valueRs.MoveNext
'    Wend
'    valueRs.Close
'    Set valueRs = Nothing
'    'MsgBox allValueStr
'    getValuesById = allValueStr
'End Function
'Public Function getExportValuesById(id As String)
'    Dim valueRs As ADODB.Recordset '���汨������������
'    Dim rowNum As String
'    Dim colName As String
'    Dim value As String
'
'    allValueStr = ""
'    Call check_condatabase
'    sql = "select row_num,col_num,value from t_baobiao_value where bb_content_id = '" & id & "' order by col_num,row_num"
'    Set valueRs = cn.Execute(sql)
'
'    While Not valueRs.EOF
'                rowNum = CStr(valueRs("row_num"))
'            colNum = CStr(valueRs("col_num"))
'            value = valueRs("value")
'
'            allValueStr = allValueStr + value & ","  '����Ϊ���� Random �� Binary ���򿪵��ļ���ֱ�����һ��ִ�е� Get ����޷����������ļ�¼ʱ��EOF ������ False��
'            If valueRs.EOF = False Then
'
'            valueRs.MoveNext
'            End If
'
'
'    Wend
'
'
'    valueRs.Close
'    Set valueRs = Nothing
'    MsgBox allValueStr
'    allValueStr = Mid(allValueStr, 1, Len(allValueStr) - 1)
'    MsgBox allValueStr
'    getExportValuesById = allValueStr
'End Function
'
'Private Sub loadData_Click()
'If Me.combox_nsrbm.Text = "" Then
'   MsgBox "���ȵ�����˰����Ϣ��"
'Exit Sub
'End If
'
'If baobiaoEditBj = "0" Then
'   MsgBox "����ѡ����౨���б�"
'   Exit Sub
'End If
'
'historyBaobiao_form.Show
'historyBaobiao_form.date_list.Clear
'historyBaobiao_form.baobiao_name = Me.lable_bb_name.Caption
'
'loadDateList
'
'End Sub
''����historyBaobiao_form��
'Private Sub loadDateList()
'
'Dim dateRs As ADODB.Recordset  '���汨�������ڵĽ����
'Dim sql As String
'Dim version As String
'version = Me.lable_version.Caption
'If version = "�°�" Then
'    version = "1"
'Else
'    version = "0"
'End If
'
'
'
'Call check_condatabase
'sql = "select bb_between_date from t_baobiao_content where user_name = '" & userName & "' and baobiao_name = '" & Me.lable_bb_name.Caption & "' and nsrbm = '" & Me.combox_nsrbm.Text & "' and version = '" & version & "'"
''MsgBox sql
'Set dateRs = cn.Execute(sql)
'While Not dateRs.EOF
'    If Trim(dateRs("bb_between_date")) <> "" Then
'         historyBaobiao_form.date_list.AddItem dateRs("bb_between_date")
'    End If
'    dateRs.MoveNext
'Wend
'End Sub
'
'
'Private Sub qi_info_frame_DragDrop(Source As Control, X As Single, Y As Single)
'
'End Sub
'
'Private Sub save_Click()
'Dim mes As String
'mes = option_validate
'If mes <> "" Then
'    MsgBox mes
'    Exit Sub
'End If
'
'
'
'
'Dim bb_content_id As String
'
'Dim nsrbm As String
'Dim baobiaoName As String
'Dim version As String  '�汾
'Dim jd As String   '����
'Dim betweenTime As String  '���ñ��
'
'nsrbm = combox_nsrbm.Text
'baobiaoName = lable_bb_name.Caption
'version = lable_version.Caption
'If version = "�°�" Then
'    version = "1"
'ElseIf version = "�ɰ�" Then
'    version = "0"
'End If
'betweenTime = CB_StartYear.Text & CB_StartMonth.Text & "-" & CB_EndYear.Text & CB_EndMonth.Text
'
'bb_content_id = saveBaobiao("0", nsrbm, baobiaoName, version, betweenTime) '���汨��Ļ�����Ϣ  �û���  ��˰�˱���   ����  �汾   ��  ������������������ID
''bb_content_id = CStr(bb_content_id)
''MsgBox bb_content_id
'deleteValue (bb_content_id)   'ÿ�α��涼��ռ�¼  ȫ�����²���
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
'       ' MsgBox valuePath
'        If Trim(valuePath) <> "" Then
'
'          '  MsgBox "sheet1��Ԫ���ַ��" & valuePath   valuePath�����ݿ���ȡ
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
'             ' MsgBox row & "   " & col & "   " & value
'              saveValue row, col, value, bb_content_id
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
'Public Sub saveValue(row As Integer, col As Integer, value As String, bb_content_id As String)
'
'        Dim valueRs As ADODB.Recordset
'        sql = "select * from t_baobiao_value"
'        Set valueRs = New ADODB.Recordset
'        Set valueRs.ActiveConnection = cn
'        valueRs.LockType = adLockOptimistic
'        valueRs.CursorType = adOpenKeyset
'        valueRs.Open sql
'
'        valueRs.AddNew '��ӱ�����Ϣ
'        valueRs("bb_content_id") = bb_content_id
'        valueRs("row_num") = row
'        valueRs("col_num") = col
'        valueRs("value") = value
'        valueRs.Update
'        valueRs.Close
'        Set valueRs = Nothing
'
'End Sub
'Public Sub deleteValue(bb_content_id As String)
'
'    Dim deleteSql As String
'    Call check_condatabase
'    deleteSql = "delete from t_baobiao_value where bb_content_id = '" & bb_content_id & "'"
'    cn.Execute (deleteSql)
'
'
'
'
'End Sub
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
'    F1Book1.CopyDataToArray sheetIndex, exlQsdX, exlQsdY, exlZdX, exlZdY, True, copyArray
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
'Set cellformat = F1Book1.CreateNewCellFormat
'With cellformat
'    .FontColor = vbBlack
'End With
'
'F1Book1.Sheet = sheetIndex
'
'F1Book1.SetActiveCell array2(0), array2(1)
'
'F1Book1.SetCellFormat cellformat
'
'    '-------------------------
'
'    '4 Ϊ����ֵ
'    F1Book1.EntrySRC(sheetIndex, array2(0), array2(1)) = array2(2)
'
'Next i
'
''isShowData = False
'End Function
'
'
Private Sub export_data_Click()

End Sub

