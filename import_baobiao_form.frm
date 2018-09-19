VERSION 5.00
Object = "{B0475000-7740-11D1-BDC3-0020AF9F8E6E}#6.1#0"; "TTF16.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form import_baobiao_form 
   Caption         =   "������     ��һ�ε���������"
   ClientHeight    =   2610
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7455
   LinkTopic       =   "Form1"
   ScaleHeight     =   2610
   ScaleWidth      =   7455
   StartUpPosition =   3  '����ȱʡ
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   2625
      Left            =   0
      ScaleHeight     =   2565
      ScaleWidth      =   7425
      TabIndex        =   0
      Top             =   0
      Width           =   7485
      Begin VB.TextBox text_baobiao_path 
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
         TabIndex        =   3
         Top             =   480
         Width           =   3675
      End
      Begin VB.CommandButton choose_file 
         Caption         =   "ѡ��..."
         Height          =   495
         Left            =   5520
         TabIndex        =   2
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton b_import_baobiao 
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
         TabIndex        =   1
         Top             =   1560
         Width           =   1695
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������ƣ�"
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
         Left            =   120
         TabIndex        =   4
         Top             =   600
         Width           =   1275
      End
   End
   Begin TTF160Ctl.F1Book F1Hidden 
      Height          =   2175
      Left            =   1200
      TabIndex        =   5
      Top             =   0
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   3836
      _0              =   $"import_baobiao_form.frx":0000
      _1              =   $"import_baobiao_form.frx":0409
      _2              =   $"import_baobiao_form.frx":0812
      _3              =   $"import_baobiao_form.frx":0C1B
      _4              =   $"import_baobiao_form.frx":1024
      _count          =   5
      _ver            =   2
   End
   Begin MSComDlg.CommonDialog CommonDia 
      Left            =   480
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "import_baobiao_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public baobiaoPath As String   '�����ļ��������ַ���  ����ÿո�ֿ�
Public importSuccess As Boolean



Private Sub b_import_baobiao_Click()
Dim pathArray
Dim xlsPath As String

If baobiaoPath <> "" Then
    pathArray = Split(baobiaoPath, ",")
    
    If UBound(pathArray) = 0 Then
    xlsPath = baobiaoPath
    import_baobiao (xlsPath)
    Else
    For i = LBound(pathArray) To UBound(pathArray)
        xlsPath = pathArray(i)
        import_baobiao (xlsPath)
    Next
    End If
    'MsgBox "����ɹ�!"
    Unload Me
    
    'operate_form.loadBaobiao
Else
    MsgBox "��ѡ���ļ���"
End If
End Sub
Private Sub Cmd_Version_Save_Click()
    
    If Combo_Version.Text <> "" And Combo_Small.Text <> "" Then
    
    ' Me.Combo_Version.Text  me.Combo_Small.Text  Me.Label_Nsrbm.Caption
    saveNsrVersion Combo_Version.Text, Label_Nsrbm.Caption
    MsgBox "����ɹ���"
    Unload Me
    Unload main_form
    
    main_form.Show
    Else
        MsgBox "����汾�Ͱ汾�ű���ѡ��"
    End If

    
    
End Sub

Private Sub import_baobiao(xlsPath As String)
        Dim xlsName As String
        xlsName = Mid$(xlsPath, InStrRev(xlsPath, "\") + 1)
        '�ж�ָ��·�����Ƿ��Ѿ�����ͬ���ļ�
        If Dir(App.Path & "\" & xlsName) = "" Then
            'Set fso = CreateObject("Scripting.FileSystemObject")
            FileCopy xlsPath, App.Path & "\" & xlsName  '���Ƶ���·��
            saveBaobiaoInfo (xlsName)     '�������ݱ�
        Else
        MsgBox xlsName & "�˱����Ѿ�����!"
        End If
        
End Sub
Private Sub saveBaobiaoInfo(xlsName As String)
getUrl App.Path & "\" & xlsName, Me.F1Hidden
Dim versionid As String
Dim baobiaoZl As String
Dim smallId As String
'Dim baobiao_name As String
Dim xybj As String
Dim versionIdStr As String
versionIdStr = getData(5000, 1, 5000, 1, Me.F1Hidden)
versionid = getThirdValue(versionIdStr)
If versionid = "" Then
   MsgBox "����汾��Ϣ��ʽ���ԣ�"
   Exit Sub
End If
Dim smallIdStr As String
smallIdStr = getData(5001, 1, 5001, 1, Me.F1Hidden)
smallId = getThirdValue(smallIdStr)
Dim baobiaoZlStr As String
baobiaoZlStr = getData(5002, 1, 5002, 1, Me.F1Hidden)
baobiaoZl = getThirdValue(baobiaoZlStr)
Dim sql As String
Dim baobiaoInfoRs As ADODB.Recordset
Dim baobiaoNameRs As ADODB.Recordset

Dim version As String  '�汾
xybj = "1"  'ѡ�ñ��  ����ʱĬ��Ϊ1  ����

'���ݿ����
Call check_condatabase
sql = "select baobiao_name from t_baobiao where baobiao_name = '" & xlsName & "'"
Set baobiaoNameRs = cn.Execute(sql)
If Not baobiaoNameRs.EOF Then
    MsgBox "�Բ��𣬴˱����Ѿ������ˣ�"
    Exit Sub
End If
baobiaoNameRs.Close
Set baobiaoNameRs = Nothing
  
sql = "select * from t_baobiao"
Set baobiaoInfoRs = New ADODB.Recordset
Set baobiaoInfoRs.ActiveConnection = cn
baobiaoInfoRs.LockType = adLockOptimistic
baobiaoInfoRs.CursorType = adOpenKeyset
baobiaoInfoRs.Open sql
  
baobiaoInfoRs.AddNew '��ӱ�����Ϣ
baobiaoInfoRs("baobiao_name") = xlsName
baobiaoInfoRs("version_id") = versionid
baobiaoInfoRs("xybj") = xybj
baobiaoInfoRs("small_id") = smallId
baobiaoInfoRs("baobiao_zl") = baobiaoZl
baobiaoInfoRs.Update
baobiaoInfoRs.Close
Set baobiaoInfoRs = Nothing
MsgBox xlsName & "����ɹ���"
End Sub
Private Sub insertBaobiao(xlsName As String)
xlsName = Mid(xlsName, 1, Len(xlsName) - 4)

Dim sql As String
Dim baobiaoInfoRs As ADODB.Recordset
Dim baobiaoNameRs As ADODB.Recordset

Dim version As String  '�汾
Dim bj As String  '���ñ��

version = "1"
bj = "1"

'���ݿ����
Call check_condatabase
sql = "select baobiao_name from t_baobiao where baobiao_name = '" & xlsName & "'"
Set baobiaoNameRs = cn.Execute(sql)
If Not baobiaoNameRs.EOF Then
    MsgBox "�Բ��𣬴˱����Ѿ������ˣ�"
    Exit Sub
End If
baobiaoNameRs.Close
Set baobiaoNameRs = Nothing
  
sql = "select * from t_baobiao"
Set baobiaoInfoRs = New ADODB.Recordset
Set baobiaoInfoRs.ActiveConnection = cn
baobiaoInfoRs.LockType = adLockOptimistic
baobiaoInfoRs.CursorType = adOpenKeyset
baobiaoInfoRs.Open sql
  
baobiaoInfoRs.AddNew '��ӱ�����Ϣ
baobiaoInfoRs("baobiao_name") = xlsName
baobiaoInfoRs("version") = version
baobiaoInfoRs("bj") = bj
baobiaoInfoRs.Update
baobiaoInfoRs.Close
Set baobiaoInfoRs = Nothing

End Sub

Private Sub choose_file_Click()
On Error GoTo errpro
Dim i As Integer, xlsName As String, FileNames As String
  baobiaoPath = ""

  CommonDia.Filter = "All   Excel   Files   (*.xls)|*.xls|All   files   (*.*)|*.*"
  CommonDia.InitDir = App.Path
  CommonDia.DialogTitle = "��ѡ��Ҫ����ı���"

  CommonDia.Flags = cdlOFNAllowMultiselect Or cdlOFNExplorer        '  ����ǹؼ�

  'CommonDia.Action = 1
  CommonDia.ShowOpen
  FileNames = CommonDia.fileName
  'MsgBox FileNames
  If FileNames <> "" Then
    a = Split(Trim(FileNames), Chr(0))
    If UBound(a) = 0 Then
      baobiaoPath = FileNames
    Else
       Dim filePath As String
       filePath = a(LBound(a))
      ' MsgBox LBound(a) + 1 & "     " & UBound(a)
       For i = LBound(a) + 1 To UBound(a)
            baobiaoPath = baobiaoPath & filePath & "\" & a(i) & ","
       Next
       baobiaoPath = Left(baobiaoPath, Len(baobiaoPath) - 1)
    End If
 End If
 
 ' MsgBox baobiaoPath

  text_baobiao_path.Text = baobiaoPath
errpro:
' MsgBox "��û��ѡ���κ��ļ����ļ������ڻ��ļ������ϡ�", vbCritical, "ѡ�����"
End Sub
    
  Function GetLeftWords(s As String, ByVal Ch As String) As String
          Dim i     As Long
          i = InStr(s, Ch)
          If i > 0 Then
                  GetLeftWords = Left(s, i - 1)
                  s = Mid(s, i + Len(Ch))
          Else
                  GetLeftWords = s
                  s = vbNullString
          End If
  End Function

