VERSION 5.00
Object = "{B0475000-7740-11D1-BDC3-0020AF9F8E6E}#6.1#0"; "TTF16.ocx"
Begin VB.Form main_form 
   Caption         =   "报表录入主界面"
   ClientHeight    =   10875
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   15240
   Icon            =   "main_form.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   10875
   ScaleWidth      =   15240
   WindowState     =   2  'Maximized
   Begin VB.CommandButton save_as 
      Caption         =   "另存为"
      Height          =   375
      Left            =   12840
      TabIndex        =   22
      Top             =   1560
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "粘贴&D"
      Height          =   375
      Left            =   8520
      TabIndex        =   21
      Top             =   1560
      Width           =   855
   End
   Begin VB.Frame Frame_Nsrxx 
      Caption         =   "纳税人信息"
      Height          =   1095
      Left            =   120
      TabIndex        =   16
      Top             =   120
      Width           =   15735
      Begin VB.ComboBox Combo_Nsrbm 
         Height          =   300
         Left            =   1320
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   480
         Width           =   1935
      End
      Begin VB.Label Label_Nsrmc_Value 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7800
         TabIndex        =   20
         Top             =   480
         Width           =   5775
      End
      Begin VB.Label Label_Nsrbm 
         Caption         =   "纳税人编码："
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   480
         Width           =   1095
      End
      Begin VB.Label Label_Nsrmc 
         Caption         =   "纳税人名称："
         Height          =   255
         Left            =   6600
         TabIndex        =   17
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.CommandButton cmd_export 
      Caption         =   "导出"
      Height          =   375
      Left            =   14040
      TabIndex        =   15
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton cmd_yanzheng 
      Caption         =   "验证"
      Height          =   375
      Left            =   11760
      TabIndex        =   14
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Cmd_Clear 
      Caption         =   "清空"
      Height          =   375
      Left            =   10680
      TabIndex        =   10
      Top             =   1560
      Width           =   855
   End
   Begin VB.CommandButton Cmd_Save 
      Caption         =   "保存"
      Height          =   375
      Left            =   9600
      TabIndex        =   9
      Top             =   1560
      Width           =   855
   End
   Begin TTF160Ctl.F1Book F1Book1 
      Height          =   7485
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   15735
      _ExtentX        =   27755
      _ExtentY        =   13203
      _0              =   $"main_form.frx":16AC2
      _1              =   $"main_form.frx":16ECB
      _2              =   $"main_form.frx":172D4
      _3              =   $"main_form.frx":176DE
      _4              =   $"main_form.frx":17AE9
      _count          =   5
      _ver            =   2
   End
   Begin VB.Frame Frame_Baobiao_Zl 
      Caption         =   "报表信息"
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   1320
      Width           =   8295
      Begin VB.CommandButton Cmd_init_data 
         Caption         =   "加载上期数据"
         Height          =   375
         Left            =   3720
         TabIndex        =   13
         Top             =   240
         Width           =   1695
      End
      Begin VB.ComboBox CB_Season 
         Height          =   300
         Left            =   2280
         TabIndex        =   5
         Top             =   360
         Width           =   855
      End
      Begin VB.ComboBox CB_Year 
         Height          =   300
         ItemData        =   "main_form.frx":17DAE
         Left            =   1200
         List            =   "main_form.frx":17DB0
         TabIndex        =   4
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label_SmallVersionID 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7680
         TabIndex        =   12
         Top             =   360
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Label Label1 
         Caption         =   "版本号："
         Height          =   255
         Left            =   6960
         TabIndex        =   11
         Top             =   360
         Visible         =   0   'False
         Width           =   735
      End
      Begin VB.Label Label_Jd 
         Caption         =   "月"
         Height          =   255
         Left            =   3240
         TabIndex        =   7
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label7 
         Caption         =   "年"
         Height          =   255
         Left            =   2040
         TabIndex        =   6
         Top             =   360
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "报表所属期："
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label_Bb_Value 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6600
         TabIndex        =   2
         Top             =   360
         Width           =   1095
      End
      Begin VB.Label Label_Bb 
         Caption         =   "版本："
         Height          =   255
         Left            =   5760
         TabIndex        =   1
         Top             =   360
         Width           =   615
      End
   End
End
Attribute VB_Name = "main_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub C_Click()

End Sub

Private Sub CB_Season_Click()
Dim dateYear As String
Dim dateSeason As String
Dim ver_name As String
Dim versionStr As String
Dim versionID As String

Dim allValueStr As String


Dim bb_content_id As String
Dim sheetName As String
Dim rs As ADODB.Recordset



  
 
  Dim nsrbm As String
  nsrbm = main_form.Combo_Nsrbm.Text
  dateYear = main_form.CB_Year.Text
  dateSeason = main_form.CB_Season.Text
  dateSeason = date_change(dateSeason)
   If dateYear <> "" And dateSeason <> "" Then
   
   
     
      versionID = get_version(nsrbm, dateYear, dateSeason)
      If versionID = "" Then
      loadBaobiao
      Exit Sub
      
      End If
      baobiaoName = get_baobiao_name(versionID)
     ' MsgBox baobiaoName
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
       ' dateSeason = change_date(dateSeason)
       ' main_form.CB_Year.Text = dateYear
       ' main_form.CB_Season.Text = dateSeason
        main_form.F1Book1.Sheet = 1
        If hy = "1" Or hy = "2" Then
        
        main_form.F1Book1.ObjValue(pid3) = main_form.F1Book1.EntryRC(3, 3)
        Else
        
        End If
        
          main_form.F1Book1.ObjValue(pid) = main_form.F1Book1.EntryRC(9, 3)
           main_form.F1Book1.ObjValue(pid2) = main_form.F1Book1.EntryRC(10, 3)
  Else
  End If
End Sub


Private Sub CB_Year_Click()
Dim dateYear As String
Dim dateSeason As String
Dim ver_name As String
Dim versionStr As String
Dim versionID As String

Dim allValueStr As String


Dim bb_content_id As String
Dim sheetName As String
Dim rs As ADODB.Recordset



  
 
  Dim nsrbm As String
  nsrbm = main_form.Combo_Nsrbm.Text
  dateYear = main_form.CB_Year.Text
  dateSeason = main_form.CB_Season.Text
  dateSeason = date_change(dateSeason)
   If dateYear <> "" And dateSeason <> "" Then
   
   
     
      versionID = get_version(nsrbm, dateYear, dateSeason)
      If versionID = "" Then
      loadBaobiao
      
      Exit Sub
      
      End If
      baobiaoName = get_baobiao_name(versionID)
     ' MsgBox baobiaoName
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
       ' dateSeason = change_date(dateSeason)
       ' main_form.CB_Year.Text = dateYear
       ' main_form.CB_Season.Text = dateSeason
        main_form.F1Book1.Sheet = 1
        If hy = "1" Or hy = "2" Then
        
        main_form.F1Book1.ObjValue(pid3) = main_form.F1Book1.EntryRC(3, 3)
        Else
        
        End If
          main_form.F1Book1.ObjValue(pid) = main_form.F1Book1.EntryRC(9, 3)
           main_form.F1Book1.ObjValue(pid2) = main_form.F1Book1.EntryRC(10, 3)
  Else
  End If
End Sub



Private Sub Cmd_Clear_Click()
Dim mes As String
mes = option_validate
If mes <> "" Then
    MsgBox mes
    Exit Sub
End If
loadbaobiao1 (main_form.Label_Bb_Value.Caption)

End Sub

Private Sub Cmd_Export_Click()
Unload Form_Excel
Form_Excel.Show
End Sub

Private Sub Cmd_init_data_Click()
Dim mes As String
mes = option_validate
If mes <> "" Then
    MsgBox mes
    Exit Sub
End If

Dim bb_content_id As String '报表信息ID
Dim nsrbm As String
Dim baobiaoName As String
Dim version As String  '版本
'Dim smallVersionID As String  '小版本号
Dim dateYear As String  '年
Dim dateSeason As String  '季度
loadBaobiao
Dim initDateYear As String
Dim initDateSeason As String

Dim bbContentRs As ADODB.Recordset
Dim id As String
id = "0"

nsrbm = Me.Combo_Nsrbm.Text
version = Me.Label_Bb_Value.Caption
'smallVersionID = Me.Label_SmallVersionID.Caption
dateYear = Me.CB_Year.Text
dateSeason = date_change(Me.CB_Season.Text)

Dim versionID As String
versionID = getVersionID(version)
 
Dim initBj As String   '初始化标记   从本sheet页中的5050行1列取

Dim sheetName As String
Dim sheetNum As Integer
sheetNum = F1Book1.NumSheets
'sheetNum = 1
For i = 1 To sheetNum    '循环每个sheet页
   sheetName = F1Book1.sheetName(i)
   
   initBj = getData4NameSheet(sheetName, 5050, 1, 5050, 1, Me.F1Book1)
   initBj = Mid(initBj, 8, 1)
   'MsgBox sheetName & initBj
   If initBj = "0" Then
        initDateYear = ""
        initDateSeason = ""
   ElseIf initBj = "1" Then   '加载年初余额   即上年第4季度的期末余额
        initDateYear = CStr(CInt(dateYear) - 1)
        initDateSeason = "4"
        'MsgBox initDateYear & "   " & initDateSeason
   ElseIf initBj = "2" Then   '加载上期金额   即上年同期的本期金额
        
        initDateYear = CStr(CInt(dateYear) - 1)
        initDateSeason = dateSeason
        'MsgBox initDateYear & "   " & initDateSeason
   ElseIf initBj = "3" Then
        initDateYear = CStr(CInt(dateYear) - 1)
        initDateSeason = dateSeason
        jz_bj = 1 '表示加载的是旧版本的利润表
        
   ElseIf initBj = "4" Then
          If dateSeason = "1" Then
          
          MsgBox ("经营信息表不能跨年加载！")
          GoTo xia
          
          
            End If
            
           initDateYear = dateYear
           initDateSeason = CStr(CInt(dateSeason) - 1)
           
            
           jz_bj = 2 '表示加载的是经营信息表
           
        
   End If
   
    If initDateYear <> "" And initDateSeason <> "" Then
       
        Call check_condatabase
        sql = "select id from t_baobiao_content where nsrbm = '" & nsrbm & "'and user_name = '" & username & "' and baobiao_name='" & sheetName & "' and version = '" & versionID & "'  and date_year = '" & initDateYear & "' and date_season = '" & initDateSeason & "'"
        'MsgBox sql
        Set bbContentRs = cn.Execute(sql)
        
        '    If Trim(nsrRs("nsrbm")) <> "" Then
        '    Me.combox_nsrbm.AddItem nsrRs("nsrbm")
        '    End If
        '    nsrRs.MoveNext
        'Wend
        If Not bbContentRs.EOF Then
            id = CStr(bbContentRs("id"))
            'MsgBox "根据条件查询得到的ID：" & id
        Else
           id = "0"
        End If
        bbContentRs.Close
        Set bbContentRs = Nothing
        
        'MsgBox "根据此ID获取所有值：" & id
        If id <> "0" Then
               Dim allValueStr As String
               allValueStr = getInitValuesById(sheetName, id)
               'MsgBox allValueStr
               showData4NameSheet sheetName, allValueStr, Me.F1Book1
        Else
           '   MsgBox "没有初始化数据！"
        End If
    End If
    
xia:
Next i
MsgBox "上期数据加载完毕！"

        Me.F1Book1.MaxRow = 70
jz_bj = 0 '释放标记
End Sub

Public Function getInitValuesById(sheetName As String, id As String)
    Dim valueRs As ADODB.Recordset '保存报表输入项结果集
    Dim rownum As String
    Dim colnum As String
    Dim value As String
    Dim allValueStr As String
    
    Dim initColNum As String
    Dim initColNumArray
    
    Dim colNumArray
    Dim fromColNum As String
    Dim toColNum As String
    'initColNum = getData4IndexSheet(2, 4, 1, 4, 1)
    initColNum = getData4NameSheet(sheetName, 5051, 1, 5051, 1, Me.F1Book1)
    initColNum = Mid(initColNum, 8, Len(initColNum) - 8)
    'MsgBox initColNum
    initColNumArray = Split(initColNum, ",")
    'MsgBox initColNumArray(0) & "  " & initColNumArray(1)
    allValueStr = ""
    If jz_bj = 1 Then
    Call check_condatabase
    sql = "select row_num,col_num,value from t_baobiao_value where bb_content_id = '" & id & "'and row_num>22"
    Set valueRs = cn.Execute(sql)

    While Not valueRs.EOF
       
        rownum = CStr(valueRs("row_num"))
        colnum = CStr(valueRs("col_num"))
        value = valueRs("value")
        
        For i = LBound(initColNumArray) To UBound(initColNumArray)
            colNumArray = Split(initColNumArray(i), "-")
            fromColNum = colNumArray(0)
            toColNum = colNumArray(1)
            If colnum = fromColNum Then
                allValueStr = allValueStr + rownum & "," & toColNum & "," & value & ";"
            End If
        Next
        
        
        valueRs.MoveNext
    Wend
    valueRs.Close
    Set valueRs = Nothing
    jz_bj = 0
    
    End If
    
    
    
    
    
     If jz_bj = 2 Then
    
    
     Call check_condatabase
    sql = "select row_num,col_num,value from t_baobiao_value where bb_content_id = '" & id & "'  "
    Set valueRs = cn.Execute(sql)

    While Not valueRs.EOF
       
        rownum = CStr(valueRs("row_num"))
        colnum = CStr(valueRs("col_num"))
        value = valueRs("value")
        
        
        
        If rownum = "3" Then
        
        If hy = "1" Or hy = "2" Then
        
       main_form.F1Book1.ObjValue(pid3) = value
       Else
       End If
       
       End If
       If rownum = "9" Then
       
       
        main_form.F1Book1.ObjValue(pid) = value
       
       End If
       If rownum = "10" Then
       
       main_form.F1Book1.ObjValue(pid2) = value
       End If
       If rownum = "4" Then
       
                allValueStr = rownum & "," & colnum & "," & value & ";"
    End If
    
        
        
        
        valueRs.MoveNext
    Wend
 
    valueRs.Close
    Set valueRs = Nothing
    jz_bj = 0
 End If
 
 
 
    
    
    
    
    
    
    
    
        If jz_bj <> 2 And jz_bj <> 1 Then
    
    Call check_condatabase
    sql = "select row_num,col_num,value from t_baobiao_value where bb_content_id = '" & id & "'"
    Set valueRs = cn.Execute(sql)

    While Not valueRs.EOF

        rownum = CStr(valueRs("row_num"))
        colnum = CStr(valueRs("col_num"))
        value = valueRs("value")
        
        For i = LBound(initColNumArray) To UBound(initColNumArray)
            colNumArray = Split(initColNumArray(i), "-")
            fromColNum = colNumArray(0)
            toColNum = colNumArray(1)
            If colnum = fromColNum Then
                allValueStr = allValueStr + rownum & "," & toColNum & "," & value & ";"
            End If
        Next
        
        
        valueRs.MoveNext
    Wend
    valueRs.Close
    Set valueRs = Nothing
    'MsgBox allValueStr
    End If
    getInitValuesById = allValueStr
End Function
Private Sub Cmd_Save_Click()

'Me.F1Book1.WriteEx "D:\aaa.xls", 11
'
'F1Book1.Write
'Exit Sub
'main_form.F1Book1.SheetSelected(1) = True
F1Book1.Sheet = 1

If hy = "1" Or hy = "2" Then


If main_form.F1Book1.ObjValue(pid3) = -1 Then
MsgBox ("主营产品必须选择")
Exit Sub
Else
'yy = main_form.F1Book1.ObjItem(pid2, main_form.F1Book1.ObjValue(pid2))
zz = main_form.F1Book1.ObjValue(pid3)
End If

End If

If main_form.F1Book1.ObjValue(pid) = -1 Then
MsgBox ("营业状态必须填写")
Exit Sub
Else
  
'xx = main_form.F1Book1.ObjItem(pid, main_form.F1Book1.ObjValue(pid))
xx = main_form.F1Book1.ObjValue(pid)


End If
If main_form.F1Book1.ObjValue(pid2) = -1 Then
MsgBox ("征收机构必须填写")
Exit Sub
Else
'yy = main_form.F1Book1.ObjItem(pid2, main_form.F1Book1.ObjValue(pid2))
yy = main_form.F1Book1.ObjValue(pid2)
End If




If hy = "1" Or hy = "2" Then

If F1Book1.EntryRC(4, 3) = "" Or F1Book1.EntryRC(5, 3) = "" Or F1Book1.EntryRC(6, 3) = "" Or F1Book1.EntryRC(7, 3) = "" Or F1Book1.EntryRC(8, 3) = "" Or F1Book1.EntryRC(11, 3) = "" Or F1Book1.EntryRC(12, 3) = "" Or F1Book1.EntryRC(13, 3) = "" Then

MsgBox ("基本信息表中蓝色部分必须填写")

Exit Sub

Else

End If

End If


If hy = "3" Then

If F1Book1.EntryRC(11, 3) = "" Or F1Book1.EntryRC(12, 3) = "" Or F1Book1.EntryRC(13, 3) = "" Then

MsgBox ("基本信息表中蓝色部分必须填写")

Exit Sub

Else

End If

 End If




Dim mes As String
mes = option_validate
If mes <> "" Then
    MsgBox mes
    Exit Sub
End If

           F1Book1.EntryRC(3, 3) = zz
           F1Book1.EntryRC(9, 3) = xx
           F1Book1.EntryRC(10, 3) = yy
          

Dim bb_content_id As String '报表信息ID
Dim nsrbm As String
Dim baobiaoName As String
Dim version As String  '版本
Dim smallVersionID As String  '小版本号
Dim dateYear As String  '年
Dim dateSeason As String  '季度



nsrbm = Me.Combo_Nsrbm.Text
version = Me.Label_Bb_Value.Caption
'smallVersionID = Me.Label_SmallVersionID.Caption
dateYear = Me.CB_Year.Text



dateSeason = date_change(Me.CB_Season.Text)

Dim versionID As String
versionID = getVersionID(version)

Dim saveFlag As Boolean
saveFlag = validate_version(nsrbm, dateYear, dateSeason)

Dim sheetName As String
Dim sheetNum As Integer
sheetNum = F1Book1.NumSheets


If saveFlag Then    '本纳税人本季度存在保存了的报表
    Dim bb_id As String
    bb_id = getBaobiaoID(nsrbm, "", versionID, dateYear, dateSeason)
    If bb_id <> "0" Then
         For i = 1 To sheetNum
         
           
           sheetName = F1Book1.sheetName(i)
           'saveBaobiaoByName sheetName
           '查看当前纳税人的当期是否存在已保存数据！若有 不能再保存别的版本的数据
      
            
            bb_content_id = saveBbInfo(nsrbm, sheetName, version, dateYear, dateSeason) '保存报表的基本信息  用户名  纳税人编码   表名  版本   等  并返回主键报表内容ID
            ' MsgBox bb_content_id
            If bb_content_id <> "0" And Val(bb_content_id) > 0 Then
         
            
             
                saveBaobiaoValues sheetName, bb_content_id
               
            Else
    
            End If
            
     
         Next i
         MsgBox "保存成功！"
     Else
             MsgBox "不能保存！因为本季度的已经录入，不可更改版本！"
     End If
Else
    
    For i = 1 To sheetNum
       sheetName = F1Book1.sheetName(i)
       'saveBaobiaoByName sheetName
       '查看当前纳税人的当期是否存在已保存数据！若有 不能再保存别的版本的数据
      
            
            bb_content_id = saveBbInfo(nsrbm, sheetName, version, dateYear, dateSeason) '保存报表的基本信息  用户名  纳税人编码   表名  版本   等  并返回主键报表内容ID
            ' MsgBox bb_content_id
            If bb_content_id <> "0" And Val(bb_content_id) > 0 Then
                saveBaobiaoValues sheetName, bb_content_id
            Else
    
            End If
            
       
    Next i
    
    MsgBox "保存成功！"

End If


End Sub



Public Function validate_version(nsrbm As String, dateYear As String, dateSeason As String) As Boolean    'exportbj为1  则说明是导出或导入  只查询ID   为0则说明是保存  查询ID  有则返回 无则插入后生成ID
Dim flag As Boolean
Dim rs As ADODB.Recordset
Dim sql As String
sql = "select * from t_baobiao_content where nsrbm = '" & nsrbm & "' and user_name='" & username & "' and date_year = '" & dateYear & "' and date_season = '" & dateSeason & "'"
Call check_condatabase
Set rs = cn.Execute(sql)
If Not rs.EOF Then
   flag = True
End If
rs.Close
Set rs = Nothing
validate_version = flag
End Function

Public Function saveBbInfo(nsrbm As String, baobiaoName As String, version As String, dateYear As String, dateSeason As String)    'exportbj为1  则说明是导出或导入  只查询ID   为0则说明是保存  查询ID  有则返回 无则插入后生成ID
Dim id As String 't_baobiao_content 的ID
Dim versionID As String
versionID = getVersionID(version)
id = getBaobiaoID(nsrbm, baobiaoName, versionID, dateYear, dateSeason)
'MsgBox id

Dim createTime As String
createTime = Date & "  " & Time


If id = "0" Then

'        If saveFlag Then
'            MsgBox "不能保存！因为本季度的已经录入，不可更改版本！"
'        Else
            sql = "select * from t_baobiao_content"
            Set valueRs = New ADODB.Recordset
            Set valueRs.ActiveConnection = cn
            valueRs.LockType = adLockOptimistic
            valueRs.CursorType = adOpenKeyset
            valueRs.Open sql
                  
            valueRs.AddNew '添加报表信息
            valueRs("nsrbm") = nsrbm
            valueRs("baobiao_name") = baobiaoName
            valueRs("version") = versionID
        
            valueRs("date_year") = dateYear
            valueRs("date_season") = dateSeason
            valueRs("user_name") = username
            valueRs("create_time") = createTime
            valueRs.Update
            id = CStr(valueRs("id"))
            valueRs.Close
            Set valueRs = Nothing

        
'    End If
    
End If
    
saveBbInfo = id

End Function


Private Sub saveBaobiaoValues(sheetName As String, bb_content_id As String)

''bb_content_id = CStr(bb_content_id)
''MsgBox bb_content_id
deleteValue (bb_content_id)   '每次保存都清空记录  全部重新插入
Dim dataRange As String   '保存标示输入项的范围的区域  如1，1，3，3 则表明在A1到C3的区域的每个单元格存放的都是一个输入范围。暂存每个报表的sheet2中
Dim valueRange As String   '保存报表可输入项的范围
dataRange = getData4NameSheet(sheetName, 5100, 1, 5100, 1, Me.F1Book1) '固定第二行第一列保存可输入项的所有坐标信息
'注意此处的截串  若在不同范围则截的长度不同 这个存放数据的范围有关  这里是从5100行1列开始存放，按列依次往下，所以只要超过5999行  此规则即可用
dataRange = Mid(dataRange, 8, Len(dataRange) - 8)   '  如dataRange = 5100,1,12,1,18,1  代表5100行1列的值为12，1，18，1  所以这个截串就是获得12,1,18,1
'MsgBox sheetName & dataRange

If Trim(dataRange) <> "" Then

    Dim dataParamArray
    dataParamArray = Split(dataRange, ",")
    'sheet2中保存范围区域参数的数组
    dataPathArray = Split(dataRange, ",")
    Dim param1 As Integer
    Dim param2 As Integer
    Dim param3 As Integer
    Dim param4 As Integer
    param1 = CInt(dataParamArray(0))
    param2 = CInt(dataParamArray(1))
    param3 = CInt(dataParamArray(2))
    param4 = CInt(dataParamArray(3))
    valueRange = getData4NameSheet(sheetName, param1, param2, param3, param4, Me.F1Book1)
    'MsgBox valueRange
    Dim valuePathArray  '存放sheet1中每段可输入单元格的范围
    Dim valuePath  '单元格坐标
    Dim valueStr  As String  'sheet1的可输入框的坐标和值字符串  即getData的返回值  如1,1,ahha;1,2,sdf;
    valuePathArray = Split(valueRange, ";")
    For i = LBound(valuePathArray) To UBound(valuePathArray) - 1
        valuePath = valuePathArray(i)
        '此处截串注意   通上面的截串注释同样问题
        valuePath = Mid(valuePath, 8, Len(valuePath) - 7)  'valuePath = 12,1,12,1,18,1  代表12行1列的值为12，1，18，1  所以这个截串就是获得12,1,18,1
       ' MsgBox sheetName & valuePath
       ' MsgBox valuePath
        If Trim(valuePath) <> "" Then

          '  MsgBox "sheet1单元格地址：" & valuePath   valuePath从数据库中取
            valueArray = Split(valuePath, ",")
            param1 = CInt(valueArray(0))
            param2 = CInt(valueArray(1))
            param3 = CInt(valueArray(2))
            param4 = CInt(valueArray(3))

            valueStr = getData4NameSheet(sheetName, param1, param2, param3, param4, Me.F1Book1)
            'MsgBox valueStr

           Dim allValueArray  '值数组  如  1,2,asdf;2,3,sdfad;1,3,sdfsf;
           Dim cellArray   '单元格数组  如  1,2,asdf
           Dim cell As String
           Dim row As Integer
           Dim Col As Integer
           Dim value As String

           allValueArray = Split(valueStr, ";")
           For j = LBound(allValueArray) To UBound(allValueArray) - 1
              cell = allValueArray(j)
              cellArray = Split(cell, ",")
              row = CInt(cellArray(0))
              Col = CInt(cellArray(1))
              value = cellArray(2)

             ' MsgBox row & "   " & col & "   " & value
              saveValue row, Col, value, bb_content_id
           Next j

        End If
    Next i

Else
MsgBox "此报表没有输入项范围信息，无法保存"
End If

End Sub
Public Sub saveValue(row As Integer, Col As Integer, value As String, bb_content_id As String)
    
        Dim valueRs As ADODB.Recordset
        sql = "select * from t_baobiao_value"
        Set valueRs = New ADODB.Recordset
        Set valueRs.ActiveConnection = cn
        valueRs.LockType = adLockOptimistic
        valueRs.CursorType = adOpenKeyset
        valueRs.Open sql
          
        valueRs.AddNew '添加报表信息
        valueRs("bb_content_id") = bb_content_id
        valueRs("row_num") = row
        valueRs("col_num") = Col
        valueRs("value") = value
        valueRs.Update
        valueRs.Close
        Set valueRs = Nothing
    
End Sub

Private Sub deleteValue(bb_content_id As String)
Dim deleteSql As String
Call check_condatabase
deleteSql = "delete from t_baobiao_value where bb_content_id = '" & bb_content_id & "'"
cn.Execute (deleteSql)
End Sub

Private Sub Cmd_ViewHistory_Click()


historyBaobiao_form.Show
historyBaobiao_form.date_list.Clear
historyBaobiao_form.nsrbm_value = Me.Combo_Nsrbm.Text
historyBaobiao_form.nsrmc_valeu = Me.Label_Nsrmc_Value.Caption

loadDateList
End Sub
'加载historyBaobiao_form的
Private Sub loadDateList()

Dim dateRs As ADODB.Recordset  '保存报表所属期的结果集
Dim sql As String
Dim version As String
'version = Me.lable_version.Caption
Dim betweenDate As String
Dim itemCount As Integer
Dim itemFlag As Boolean
Call check_condatabase
sql = "select * from t_baobiao_content where user_name = '" & username & "' and nsrbm = '" & Me.Combo_Nsrbm.Text & "'"
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

Private Sub cmd_yanzheng_Click()
'Load main_form
 '   If userType = "1" Then
  '      Unload Form_Export
   '     Form_Export.nsrbm_value.Caption = main_form.Combo_Nsrbm.Text
       
   ' ElseIf userType = "0" Then
    '    Export_Many_Form.Show
    'End If
    Unload Form_Export
   Form_Export.Show
    
End Sub

Private Sub Combo_Nsrbm_Click()
Dim nsrRs As ADODB.Recordset '保存纳税人的结果集
Dim sql As String

Dim nsrbmText As String
Dim nsrmcText As String
Dim nsrdzText As String
Dim zclxText As String
Dim zchyText As String

nsrbmText = Trim(Combo_Nsrbm.Text)

Call check_condatabase
If nsrbmText <> "" Then
    sql = "select * from t_nsrxx where username='" & username & "' and nsrbm ='" & nsrbmText & "'"
    Set nsrRs = cn.Execute(sql)
    If Not nsrRs.EOF Then
        nsrRs.MoveFirst
        Label_Nsrmc_Value.Caption = nsrRs("nsrmc")
       ' Label_Nsrdz_Value.Caption = nsrRs("nsrdz")
       ' Label_Zclx_Value.Caption = nsrRs("zclx")
         'Label_Hylx_Value.Caption = nsrRs("zchy")
        hy = nsrRs("zchy")
      
    End If
   
  Dim a As Integer
  Dim B As Integer
  Dim season As String
  
 
  Dim allValueStr As String
  Dim versionID As String
  Dim bb_content_id As String
  Dim sheetName As String
  Dim aRs As ADODB.Recordset
  Dim bRs As ADODB.Recordset
  Dim rs As ADODB.Recordset
  Dim versionName As String
  a = (latest_year(username, Me.Combo_Nsrbm.Text)) '通过用户名，纳税人编码 找到录入数据的最近年份
  
  B = (latest_season(a)) '年份找到最近月份
  If a <> 0 And B <> 0 Then
   loadVersion
   loadBaobiao
   Call check_condatabase
        sql = "select version from t_baobiao_content where user_name = '" & username & "' and nsrbm = '" & Me.Combo_Nsrbm.Text & "' and date_year = '" & a & "' and date_season = '" & B & "'"
        Set aRs = cn.Execute(sql)
            versionID = aRs("version")
            versionName = getVersionNameById(versionID)
            aRs.Close
         If versionName = main_form.Label_Bb_Value.Caption Then
         
         sql = "select id,baobiao_name,version from t_baobiao_content where user_name = '" & username & "' and nsrbm = '" & Me.Combo_Nsrbm.Text & "' and date_year = '" & a & "' and date_season = '" & B & "'"
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
        season = change_date(CStr(B))
        main_form.CB_Year.Text = a
        main_form.CB_Season.Text = season
        main_form.F1Book1.Sheet = 1
        
     If hy = "1" Or hy = "2" Then
     
        If main_form.F1Book1.EntryRC(9, 3) <> "" And main_form.F1Book1.EntryRC(10, 3) <> "" And main_form.F1Book1.EntryRC(3, 3) <> "" Then
        
          main_form.F1Book1.ObjValue(pid) = main_form.F1Book1.EntryRC(9, 3)
           main_form.F1Book1.ObjValue(pid2) = main_form.F1Book1.EntryRC(10, 3)
           main_form.F1Book1.ObjValue(pid3) = main_form.F1Book1.EntryRC(3, 3)
           Else
           End If
           End If
        If hy = "3" Then
        
           If main_form.F1Book1.EntryRC(9, 3) <> "" And main_form.F1Book1.EntryRC(10, 3) <> "" Then
        
          main_form.F1Book1.ObjValue(pid) = main_form.F1Book1.EntryRC(9, 3)
           main_form.F1Book1.ObjValue(pid2) = main_form.F1Book1.EntryRC(10, 3)
          
           Else
           End If
           End If
           
           
           
         Else
        ' MsgBox ("最近数据和用户选择的模板类型不一致")
        End If
        Else
          main_form.CB_Year.Text = 2009
          main_form.CB_Season.Text = "1-3"
       ' main_form.CB_Season.Text =
        loadVersion
        loadBaobiao
        End If
End If

'If nsrbmText <> "" Then
'    sql = "select baobiao_id from t_baobiao_version, where user_name='" & userName & "' and nsrbm ='" & nsrbmText & "'"
'    Set nsrRs = cn.Execute(sql)
'    If Not nsrRs.EOF Then
'        nsrRs.MoveFirst
'        Label_Nsrmc_Value.Caption = nsrRs("nsrmc")
'        Label_Nsrdz_Value.Caption = nsrRs("nsrdz")
'        Label_Zclx_Value.Caption = nsrRs("zclx")
'        Label_Hylx_Value.Caption = nsrRs("zchy")
'    End If
'End If

End Sub


Private Sub Command2_Click()
Dim dataRange As String
dataRange = getData4NameSheet("经营信息表", 5100, 1, 5100, 1, Me.F1Book1)
'MsgBox dataRange
End Sub

Private Sub Command4_Click()

End Sub

Private Sub Combo_Version_Click()

  Dim versionID As String
  Dim sql As String
  Dim versionRs As ADODB.Recordset
  Dim itemCount As Integer
  
  versionID = getVersionID(Combo_Version.Text)
  'MsgBox versionId
  'Exit Sub
  
  
  sql = "select small_id from t_baobiao where version_id = '" & versionID & "'"
  Set versionRs = cn.Execute(sql)
  
  itemCount = Combo_Small.ListCount - 1
  While itemCount >= 0
    Combo_Small.RemoveItem itemCount
    itemCount = itemCount - 1
  Wend
  
  While Not versionRs.EOF '
     Combo_Small.AddItem versionRs("small_id")

    versionRs.MoveNext
  Wend
End Sub
Public Sub loadVersionCombox()
    Combo_Version.Clear
    'Me.Combo_Small.Clear
    
    Dim versionRs As ADODB.Recordset '保存纳税人的结果集
    Dim sql As String
    
    Call check_condatabase
    sql = "select t_year_dm.year from t_year_dm,t_baobiao where t_baobiao.version_id = t_year_dm.version_id"
    Set versionRs = cn.Execute(sql)
    While Not versionRs.EOF
        If Trim(versionRs("year")) <> "" Then
       ' AddItem nsrRs("nsrbm
       Combo_Version.AddItem versionRs("year")
        End If
        versionRs.MoveNext
    Wend
    
   ' If Combo_Version.ListCount > 0 Then
   ' Combo_Version.ListIndex = 0
   ' End If
End Sub

Private Sub daoru_Click()

Dim userArray '用户信息数组

Dim nsrbmRs As ADODB.Recordset  '用户的数据库的结果集
  '用户的数据库的结果集
Dim sql As String
Dim userInfoArray
Dim nsrbm As String: nsrbm = ""  '纳税人编码
Dim NSRMC As String: NSRMC = "" '纳税人名称
Dim version As String
Dim s_version As String
Dim b_id As String
Dim bbid As String
If NSR_BM.Text <> "" And NSR_MC.Text <> "" And Combo_Version.Text <> "" And Combo_Small.Text <> "" Then
   nsrbm = NSR_BM.Text
   nsrqc = NSR_MC.Text
   version = Combo_Version.Text
  s_version = Combo_Small.Text
  Dim rs As ADODB.Recordset
  sql = "select * from t_year_dm where year = '" & version & "'"
  
  Call check_condatabase
  
  Set rs = cn.Execute(sql)
    bbid = rs("version_id")
    rs.Close
    Dim aRs As ADODB.Recordset
  sql = "select * from t_baobiao where version_id = '" & bbid & "' And small_id='" & s_version & " '"
  
  Call check_condatabase
  
  Set aRs = cn.Execute(sql)
   b_id = aRs("id")
    

Else
   MsgBox ("请输入完整信息")
End If
                '数据库操作
Call check_condatabase

sql = "select nsrbm from t_nsrxx where nsrbm = '" & nsrbm & "' and username='" & username & "' and baobi_id='" & bbid & " '"
Set nsrbmRs = cn.Execute(sql)
If Not nsrbmRs.EOF Then
   MsgBox "此纳税人信息已经导入！"
End If
nsrbmRs.Close
Set nsrbmRs = Nothing
Dim nsrxxRs As ADODB.Recordset
sql = "select * from t_nsrxx"
Set nsrxxRs = New ADODB.Recordset
Set nsrxxRs.ActiveConnection = cn
nsrxxRs.LockType = adLockOptimistic
nsrxxRs.CursorType = adOpenKeyset

nsrxxRs.Open sql
nsrxxRs.AddNew '添加纳税人信息
nsrxxRs("nsrbm") = nsrbm
nsrxxRs("nsrmc") = nsrqc
nsrxxRs("username") = username
nsrxxRs("baobi_id") = b_id
nsrxxRs.Update
nsrxxRs.Close
Set nsrxxRs = Nothing
MsgBox "录入成功！"
                      
End Sub




Private Sub Command1_Click()
 Dim s As String
 Dim linearray '剪切板中的内容
 Dim colarray '分割以后的内容
 Dim qs1array '用户选定的范围
 Dim qsarray '起始单元格
 Dim zzarray '终止单元格
 Dim linearrayLen As Integer
 
 s = Clipboard.GetText()
 If s = "" Then
 Else
 linearray = Split(s, vbCrLf)
 'colarray = Split(linearray(0), vbTab)
 qs1array = Split(Me.F1Book1.Selection, ":") '$C$3:$D$4
 qsarray = Split(qs1array(0), "$") '$C$3 qsarray(1)=C,qsarray(2)=3
  Dim X As Long
    X = change_coll(CStr(qsarray(1)))
'If UBound(qs1array) - LBound(qs1array) = 0 Then
'Else
 'zzarray = Split(qs1array(1), "$") 'zzarray(1)=D,zzarray(2)=4
 'End If
 For linearrayLen = LBound(linearray) To UBound(linearray) - 1
    colarray = Split(linearray(linearrayLen), vbTab)
    For colarraylen = LBound(colarray) To UBound(colarray)
   
    Me.F1Book1.EntryRC(CLng(qsarray(2)), X) = colarray(colarraylen)
    X = CLng(X + 1)
    Next
    X = change_coll(CStr(qsarray(1)))
    qsarray(2) = CLng((qsarray(2)) + 1)
    
 Next

End If

End Sub



Private Sub Commanss_Click()
Me.F1Book1.EditPaste
End Sub

Private Sub Commaq_Click()
Me.F1Book1.EditCopy
End Sub


Private Sub F1Book1_KeyDown(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyDelete Then
KeyCode = vbKeyBack

ElseIf Shift = vbCtrlMask And KeyCode = vbKeyC Then
Me.F1Book1.EditCopy
End If
End Sub

Private Sub F1Book1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) '显示所选范围值的和

 Dim lxarray
 Dim qs1array '用户选定的范围
 Dim qsarray '起始单元格
 Dim zzarray '终止单元格
 Dim linearrayLen As Long
 Dim linearraylen1 As Long
 Dim f1_value As Double
 Dim f2_value As Double
 Dim linearraylen2 As Long
 Dim xyarray
qs1array = Split(Me.F1Book1.Selection, ",") '$C$3:$D$4

'If UBound(qs1array) - LBound(qs1array) = 0 Then
'zz_value = F1Book1.Number
'Else
For linearrayLen = LBound(qs1array) To UBound(qs1array)
 If Len(qs1array(linearrayLen)) > 6 Then
 
 lxarray = Split(qs1array(linearrayLen), ":") '$C$3 qsarray(1)=C,qsarray(2)=3
 qsarray = Split(lxarray(0), "$")
 zzarray = Split(lxarray(1), "$")
 
 
 For linearraylen1 = qsarray(2) To zzarray(2)

  For linearraylen2 = change_coll(CStr(qsarray(1))) To change_coll(CStr(zzarray(1)))
     f1_value = f1_value + F1Book1.NumberRC(linearraylen1, linearraylen2)
   
  Next
 Next
Else
xyarray = Split(qs1array(linearrayLen), "$")
f2_value = f2_value + F1Book1.NumberRC(xyarray(2), change_coll(CStr(xyarray(1))))

End If
Next
zz_value = f1_value + f2_value

'End If
MainForm.StatusBar1.Panels(2).Text = "求和=" & zz_value
End Sub

Private Sub Form_Load()
  Dim i As Integer
  
  Me.Width = ScaleX(1024, vbPixels, vbTwips)   '设定窗体的宽度为800像素
  Me.Height = ScaleY(680, vbPixels, vbTwips)  '设定窗体的高度为680像素
  'loadVersionCombox
  loadDate  '加载年月下拉框
  loadNsrCombox   '加载纳税人信息
  'loadBaobiao   '加载报表
  'loadDate  '加载年月下拉框
  'loadVersion
  isAllowEdit True, Me.F1Book1
  Me.F1Book1.ShowEditBar = False
  Me.F1Book1.ShowTabs = F1TabsBottom
  Me.F1Book1.AllowEditHeaders = False
  
  
  
 

 ' F1Book1.ShowLockedCellsError = False '不提示锁定单元格不能修改
 
End Sub
Public Function latest_year(username As String, nsrbm As String) As Integer '通过用户名纳税人编码找到最近年
 
  Dim a As Integer

  Dim aRs As ADODB.Recordset
  Dim sql As String

  Call check_condatabase
  sql = "select * from t_baobiao_content where user_name='" & username & "' and nsrbm='" & Me.Combo_Nsrbm.Text & "'"
  Set aRs = cn.Execute(sql)
    While Not aRs.EOF
        If CInt(aRs("date_year")) > a Then
         a = CInt(aRs("date_year"))
        End If
        aRs.MoveNext
       
    Wend
    latest_year = a
End Function
Public Function latest_season(a As Integer) As Integer '通过年份找到最近月
 Dim B As Integer

  Dim aRs As ADODB.Recordset
  Dim sql As String

  Call check_condatabase
  sql = "select * from t_baobiao_content where user_name='" & username & "' and nsrbm='" & Me.Combo_Nsrbm.Text & "'and date_year='" & a & "'"
  Set aRs = cn.Execute(sql)
    While Not aRs.EOF
        If CInt(aRs("date_season")) > B Then
         B = CInt(aRs("date_season"))
        End If
        aRs.MoveNext
       
    Wend
    latest_season = B


End Function
Public Sub loadBaobiao()
Dim baobiaoUrl As String
Dim baobiaoName As String
Dim rs As ADODB.Recordset '保存纳税人的结果集
Dim yrs As ADODB.Recordset
Dim hyRs As ADODB.Recordset
Dim hymxrs As ADODB.Recordset
Dim hymc As String

Dim sql As String
Dim baobiaoID As String
Dim ver_id As String
 
baobiaoID = getbbID(Me.Combo_Nsrbm.Text)
If baobiaoID <> "" Then

    
sql = "select * from t_baobiao where id =" & CLng(baobiaoID)
    Set rs = cn.Execute(sql)
    If Not rs.EOF Then
        baobiaoName = rs("baobiao_name")
        ver_id = rs("version_id")
        sql = "select year from t_year_dm where  version_id=  '" & ver_id & "'  "
        Set yrs = cn.Execute(sql)
        If Not yrs.EOF Then
        main_form.Label_Bb_Value = yrs("year")
        End If
    
        baobiaoUrl = App.Path & "\" & baobiaoName
        getUrl baobiaoUrl, Me.F1Book1
        
  sql = "select zchy  from t_nsrxx where nsrbm= '" & Me.Combo_Nsrbm.Text & "'  and username='" & username & " '   "
  
  Set hyRs = cn.Execute(sql)
  
  hy = hyRs("zchy")
  
 ' MsgBox (hy)
 
  If hy <> "3" Then
  
  pid3 = Me.F1Book1.ObjCreate(F1ObjDropDown, 2, 2, 3, 15)
  
  sql = " select cymc from t_cy where hy='" & hy & " ' order by id"
  
  
  Set hymxrs = cn.Execute(sql)
  
  
  While Not hymxrs.EOF
  
  Me.F1Book1.ObjAddItem pid3, hymxrs("cymc")
  
  hymxrs.MoveNext
  
  Wend
  
 
  End If
  
  
 
  
  
  
       ' Me.F1Book1.ShowTabs = F1TabsOff
        'Me.F1Book1.AllowDesigner = False
          pid = Me.F1Book1.ObjCreate(F1ObjDropDown, 2, 8, 3, 15)
           Me.F1Book1.ObjAddItem pid, "扩大经营"
           Me.F1Book1.ObjAddItem pid, "正常经营"
           Me.F1Book1.ObjAddItem pid, "减产经营"
           Me.F1Book1.ObjAddItem pid, "半停产"
           Me.F1Book1.ObjAddItem pid, "停产"
           
           'Me.F1Book1.ObjItem(pid, 0)
           'Me.F1Book1.ObjValue (pid)
           pid2 = Me.F1Book1.ObjCreate(F1ObjDropDown, 2, 9, 3, 15)
           Me.F1Book1.ObjAddItem pid2, "国税"
           Me.F1Book1.ObjAddItem pid2, "地税"
           Me.F1Book1.ObjAddItem pid2, "国税地税混合"
           Me.F1Book1.ObjAddItem pid2, "不征收"
           Me.F1Book1.MaxRow = 70
           
           Dim i As Integer
           For i = 1 To F1Book1.NumSheets
           F1Book1.Sheet = i
           F1Book1.ShowLockedCellsError = False
         ' F1Book1.LaunchWorkbookDesigner = False
            Me.F1Book1.BackColor = &HFFFF00
            Me.F1Book1.MaxRow = 70
            Me.F1Book1.AllowEditHeaders = False
           Next i
           F1Book1.Sheet = 1
           
'        Dim aa As Boolean
'
'        aa = Me.F1Book1.SetValidationRule (.SetValidationRule("资产负债表!G35>资产负债表!G36", "大小判断！")
    Else
    MsgBox "数据不存在！"
    End If
End If


End Sub
Private Function getBaobiaoID(nsrbm As String, baobiaoName As String, versionID As String, dateYear As String, dateSeason As String)
Dim baobiaoRs As ADODB.Recordset
Dim valueRs As ADODB.Recordset
Dim id As String 't_baobiao_content 的ID
id = "0"  '默认为0
Call check_condatabase
If baobiaoName = "" Then
    sql = "select * from t_baobiao_content where user_name = '" & username & "' and  nsrbm = '" & nsrbm & "' and version ='" & versionID & "'  and date_year='" & dateYear & "' and date_season = '" & dateSeason & "'"
Else
    sql = "select * from t_baobiao_content where user_name = '" & username & "' and  nsrbm = '" & nsrbm & "' and baobiao_name = '" & baobiaoName & "' and version ='" & versionID & "'  and date_year='" & dateYear & "' and date_season = '" & dateSeason & "'"
End If

Set baobiaoRs = cn.Execute(sql)
If Not baobiaoRs.EOF Then
    id = CStr(baobiaoRs("id"))
    'MsgBox id
End If
baobiaoRs.Close
Set baobiaoRs = Nothing
getBaobiaoID = id
End Function
Public Sub loadNsrCombox()
    Combo_Nsrbm.Clear
    Dim nsrRs As ADODB.Recordset '保存纳税人的结果集
    Dim sql As String
    
    Call check_condatabase
    sql = "select nsrbm from t_nsrxx where username='" & username & "'"
    Set nsrRs = cn.Execute(sql)
    While Not nsrRs.EOF
        If Trim(nsrRs("nsrbm")) <> "" Then
        Me.Combo_Nsrbm.AddItem nsrRs("nsrbm")
        End If
        nsrRs.MoveNext
    Wend
    
    If Combo_Nsrbm.ListCount > 0 Then
    Me.Combo_Nsrbm.ListIndex = 0
    End If
End Sub
Public Sub loadVersion()
    'Me.Label_Bb_Value.Caption = ""
    Dim bbrs As ADODB.Recordset '保存纳税人的结果集
    Dim sql As String
    Dim baobiaoID As String
    Dim versionName As String
    'Dim smallVersionID As String
    
    baobiaoID = getbbID(Me.Combo_Nsrbm.Text)
    If baobiaoID <> "" Then
    
    sql = "select t_year_dm.year from t_baobiao,t_year_dm where t_baobiao.id =" & CLng(baobiaoID) & " and t_year_dm.version_id = t_baobiao.version_id"
    Set bbrs = cn.Execute(sql)
    If Not bbrs.EOF Then
        versionName = bbrs("year")
       ' smallVersionID = bbrs("small_id")
        'MsgBox versionName & "    " & smallVersionID
        If Me.Label_Bb_Value.Caption = versionName Then
        Else
           Me.Label_Bb_Value.Caption = bbrs("year")
           
           If Me.Label_Bb_Value.Caption <> "" Then
             loadBaobiao
           End If
        End If
    Else

    End If
    
    Else
        Me.Label_Bb_Value.Caption = ""
       ' Me.Label_SmallVersionID.Caption = ""
    End If
    
End Sub

Public Sub loadDate()
  Dim i As Integer
  For i = 2009 To 2030
    Me.CB_Year.AddItem CStr(i)
  Next i
  
    Me.CB_Season.AddItem "1-3"
    Me.CB_Season.AddItem "1-6"
    Me.CB_Season.AddItem "1-9"
    Me.CB_Season.AddItem "1-12"
      
  
End Sub

Public Function option_validate()
'Dim ok_flag As Boolean
If Me.Combo_Nsrbm.Text = "" Then
   option_validate = "请先导入纳税人信息！"
ElseIf Me.Label_Bb_Value.Caption = "" Then
   option_validate = "请先选择报表版本！"
ElseIf Trim(CB_Year.Text) = "" Or Trim(CB_Season.Text) = "" Then
     option_validate = "报表所属期不可为空！"
Else
     option_validate = ""
End If

End Function

Private Sub Form_Resize()
Me.WindowState = "2"
End Sub


Private Sub Label_Zclx_Value_Click()

End Sub

Private Sub Text1_Change()

End Sub

Private Sub save_as_Click()
Unload save_as_form
save_as_form.Show
End Sub
