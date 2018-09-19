Attribute VB_Name = "global"
Public cn As ADODB.Connection '全局的数据库连接对象
Public czry_flag As String '保存用户权限字符串
Public isxg As Boolean '是添加用户还是修改用户
Public nodename As String '保存某个用户的帐户名
Public rspsw As ADODB.Recordset '用于用户帐号密码设置的结果集合
Public database_data As String '保存导出数据表信息时的access数据库名称
Public username As String '保存登陆的用户名
Public userType As String
Public pid, pid2, pid3 As Long 'formula one 中控件返回标记
Public xx, yy, zz As String
Public jz_bj As Integer '加载上期数据的时候是否是旧版本的利润表
Public sn, sjd As Integer '提示的时候是否显示上年度
Public jdz As Integer '是否取绝对值
Public fh As Integer '是否取负
Public hy  As String '判断行业
 









Public Sub condatabase()           '创建连接到feiyong数据库的记录源                                     连接本地数据库JIMMY
    Set cn = New ADODB.Connection
      '  cn.Provider = "sqloledb"
      '  cn.Properties("Data Source").Value = "JIMMY"       建立与本地数据库的连接
      '  cn.Properties("Initial Catalog").Value = "YAOFEI"   数据库的名称
      '  cn.Properties("Integrated Security").Value = "SSPI"
      
      cn.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data source =" & App.Path & "\financialForm.mdb" '我转换的access数据库路径
      cn.Properties("Jet OLEDB:Database Password") = "niuey" 'ACCESS 密码
      cn.Open
End Sub

Public Sub check_condatabase()
If cn.State = 1 Then 'cn.State的值为1表示数据库处于连接状态
Else
   Call condatabase
End If
End Sub

Public Function date_change(sea_son As String) As String '月变季节
    If sea_son = "1-3" Then
       sea_son = "1"
ElseIf sea_son = "1-6" Then
       sea_son = "2"
ElseIf sea_son = "1-9" Then
        sea_son = "3"
ElseIf sea_son = "1-12" Then
       sea_son = "4"
ElseIf sea_son = "1-1" Then
        sea_son = "4"
 
End If
date_change = sea_son
End Function

Public Function change_date(sea_son As String) As String '季节变月
Dim sql As String
Dim rs As ADODB.Recordset
sql = "select * from t_month_dm where season_id = '" & sea_son & " '"
Call check_condatabase
Set rs = cn.Execute(sql)
change_date = rs("month_range")
rs.Close
End Function


Public Sub close_condatabase()   '关闭数据源
If cn.State = 1 Then
   cn.Close
End If
End Sub
'传入文件路径   获取txt文件的内容
Public Function getTxt(txtPath As String) As String
Dim i As Integer: i = FreeFile   '声明一个空闲文件号
Open txtPath For Input As #i
    getTxt = StrConv(InputB(LOF(i), i), vbUnicode)  ' 根据系统的缺省码页将字符串转成  中英文都行
Close #i
End Function
'传入文件路径和行号   获取txt文件的其中一行
Public Function getLine(txtPath As String, lineNum As Integer) As String
getLine = Split(getTxt(txtPath), vbCrLf)(lineNum - 1)
End Function

'fileNum = FreeFile '产生一个最佳文件号
'Open App.Path & "\user_info.ini" For Input As #fileNum  '打开文件"
'Open file_path & "\纳税人信息.txt" For Input As #fileNum  '打开文件"
'Do While Not EOF(fileNum)
'Line Input #fileNum, text_line
'array_user_info = Split(text_line, ",")
'MsgBox text_line
'Loop
'Close #fileNum   '关闭文件
'传入文件路径和行号   获取txt文件的其中一行

'获取文件中内容  行组成的数组
Public Function getLineArray(txtPath As String) As String()
Dim linearray() As String
linearray = Split(getTxt(txtPath), vbCrLf)  '按回车劈成字符串数组
getLineArray = linearray
End Function
'如 2，4，sfsdf;  根据此字符串获得sfsdf
Public Function getThirdValue(str As String) As String
    If str <> "" Then
        str = Mid(str, 1, Len(str) - 1)
        Dim tempArray
        Dim theThirdValue As String: theThirdValue = ""
        tempArray = Split(str, ",")
        If UBound(tempArray) > 1 Then
            theThirdValue = tempArray(2)
            getThirdValue = theThirdValue
        End If
    Else
       getThirdValue = ""
    End If
    
End Function

'通过version名获得
Public Function getVersionID(version_year As String) As String
  Dim versionRs As ADODB.Recordset
  Call check_condatabase
  Set versionRs = cn.Execute("select * from t_year_dm where year = '" & version_year & "'")
  If Not versionRs.EOF Then
    getVersionID = versionRs("version_id")
  Else
    getVersionID = ""
  End If
   
End Function
'通过version名获得
Public Function getVersionNameById(versionID As String) As String
  Dim versionRs As ADODB.Recordset
  'MsgBox versionID
  Call check_condatabase
  Set versionRs = cn.Execute("select * from t_year_dm where version_id = '" & versionID & "'")
  
  If Not versionRs.EOF Then
    getVersionNameById = versionRs("year")
  Else
    getVersionNameById = ""
  End If
   
End Function
Public Function getBaobiaoID(versionID As String) As String
  Dim rs As ADODB.Recordset
  Dim sql As String
  sql = "select * from t_baobiao where version_id = '" & versionID & "' "
  
  Call check_condatabase
  
  Set rs = cn.Execute(sql)
  If Not rs.EOF Then
    getBaobiaoID = rs("id")
  Else
    getBaobiaoID = ""
  End If
   
End Function

Public Function getBaobiaoID_2(versionID As String, bbzl As String) As String
  Dim rs As ADODB.Recordset
  Dim sql As String
  sql = "select * from t_baobiao where version_id = '" & versionID & "' and baobiao_zl='" & bbzl & " ' "
  
  Call check_condatabase
  
  Set rs = cn.Execute(sql)
  If Not rs.EOF Then
    getBaobiaoID_2 = rs("id")
  Else
    getBaobiaoID_2 = ""
  End If
   
End Function

Public Function getbbID(nsrbm As String) As String
    Dim rs As ADODB.Recordset '保存纳税人的结果集
    Dim sql As String
    Dim baobiaoID As String
    
    Call check_condatabase
    sql = "select baobiao_id from t_baobiao_version where user_name='" & username & "' and nsrbm = '" & nsrbm & "'"
    Set rs = cn.Execute(sql)
    If Not rs.EOF Then
        baobiaoID = rs("baobiao_id")
        getbbID = baobiaoID
    Else
    getbbID = ""
    'MsgBox "加载报表错误！请检查数据库"
    End If
   
End Function

Public Function getValuesById(id As String)
    Dim valueRs As ADODB.Recordset '保存报表输入项结果集
    Dim rownum As String
    Dim colName As String
    Dim value As String

    allValueStr = ""
    Call check_condatabase
    sql = "select row_num,col_num,value from t_baobiao_value where bb_content_id = '" & id & "'"
    Set valueRs = cn.Execute(sql)
    
    While Not valueRs.EOF
        rownum = CStr(valueRs("row_num"))
        colnum = CStr(valueRs("col_num"))
        value = valueRs("value")
        
        allValueStr = allValueStr + rownum & "," & colnum & "," & value & ";"
        valueRs.MoveNext
    Wend
    valueRs.Close
    Set valueRs = Nothing
    'MsgBox allValueStr
    getValuesById = allValueStr
End Function
Public Function getMonthRange(seasonId As String) As String
  Dim monthRs As ADODB.Recordset
  'MsgBox versionID
  Call check_condatabase
  Set monthRs = cn.Execute("select * from t_month_dm where season_id = '" & seasonId & "'")
  
  If Not monthRs.EOF Then
    getMonthRange = monthRs("month_range")
  Else
    getMonthRange = ""
  End If
   
End Function

Public Function saveNsrVersion(versionName As String, nsrbm As String, hyzl As String)
 Dim baobiaoID As String
    Dim qybj As String
    qybj = "1"   '默认为1  即可用
    
        Dim versionID As String
        versionID = getVersionID(versionName)
        baobiaoID = getBaobiaoID_2(versionID, hyzl)
        If baobiaoID <> "" Then
              
             '数据库操作
                Dim sql As String
                Dim versionRs As ADODB.Recordset
                
                Call check_condatabase
                sql = "select nsrbm from t_baobiao_version where nsrbm = '" & nsrbm & "' and user_name='" & username & "'"
                
                Set versionRs = cn.Execute(sql)
                If Not versionRs.EOF Then
                    Dim choose As Integer
                    choose = MsgBox("报表版本已被设置！是否更新", vbOKCancel)
                    If choose = 1 Then
                     sql = "delete from t_baobiao_version where nsrbm = '" & nsrbm & "' and user_name='" & username & "'"
                     Call check_condatabase
                     cn.Execute (sql)
                     saveNsrBbVersion baobiaoID, nsrbm
                    Else
                       ' Exit Sub
                    End If
                Else
                    saveNsrBbVersion baobiaoID, nsrbm
                End If
                versionRs.Close
                Set versionRs = Nothing
                
            
        Else
        MsgBox "此版本报表信息不存在！"
        End If
End Function
Public Sub saveNsrBbVersion(baobiaoID As String, nsrbm As String)
            Dim qybj As String
            qybj = "1"
            Dim valueRs As ADODB.Recordset
            sql = "select * from t_baobiao_version"
            Set valueRs = New ADODB.Recordset
            Set valueRs.ActiveConnection = cn
            valueRs.LockType = adLockOptimistic
            valueRs.CursorType = adOpenKeyset
            valueRs.Open sql
              
            valueRs.AddNew '添加报表信息
            valueRs("user_name") = username
            valueRs("nsrbm") = nsrbm
            valueRs("baobiao_id") = baobiaoID
            valueRs("qybj") = qybj
            valueRs.Update
            valueRs.Close
            Set valueRs = Nothing
           ' MsgBox "保存成功！"
End Sub
 
 

 


Public Sub updatensr_hy(hy As String, bm As String)

 Dim hyRs As ADODB.Recordset
Dim strsql As String
Dim hy_xg As String

 If hy = "工业企业" Then
 hy_xg = "1"
 ElseIf hy = "房地产企业" Then
 hy_xg = "2"
 Else
 hy_xg = "3"
 End If
 
 

Call check_condatabase
Set hyRs = New ADODB.Recordset
Set hyRs.ActiveConnection = cn
strsql = "select * from t_nsrxx   where nsrbm= '" & bm & " ' and username='" & username & "'   "
 
hyRs.Open strsql, cn, adOpenKeyset, adLockOptimistic

If Not hyRs.EOF Then
  
    hyRs.Fields("zchy") = hy_xg '是否有问题 ？？
        hyRs.Update
       
       Else
       
        End If
End Sub

Public Function getExportData(nsrbm As String, dateStr As String) As String
  Dim dateYear As String
   Dim dateSeason As String


    Dim bbValueStr As String  '一个报表的所有值
    
    Dim bb_content_id As String
    Dim sheetName As String
    Dim rs As ADODB.Recordset
    Dim fileName As String
    Dim allValueStr As String
    Dim versionID As String
    'Dim smallVersionID As String
    Dim versionName As String
    Dim monthRange As String
    
    Dim headInfo As String

    'MsgBox dateStr
    allValueStr = ""
    fileName = nsrbm & "-" & dateStr & ".txt"

        dateYear = Mid(dateStr, 1, 4)
        dateSeason = Mid(dateStr, 6, 1)
        ' MsgBox dateYear & "  " & dateSeason
            
        Call check_condatabase
        sql = "select id,baobiao_name,version,small_version_id from t_baobiao_content where user_name = '" & username & "' and nsrbm = '" & nsrbm & "' and date_year = '" & dateYear & "' and date_season = '" & dateSeason & "'"
        Set rs = cn.Execute(sql)
        While Not rs.EOF
            sheetName = rs("baobiao_name")
            bb_content_id = rs("id")
            versionID = rs("version")
            smallVersionID = rs("small_version_id")
            bbValueStr = getExportValuesById(bb_content_id)  '  1,1,sdsd;2,1,asd;1,2,dfad;2,2,dfa;
            allValueStr = allValueStr & bbValueStr
         
            rs.MoveNext
        Wend
            rs.Close
            Set rs = Nothing
            
        If allValueStr <> "" Then
        
        allValueStr = Mid(allValueStr, 1, Len(allValueStr) - 1)
        
        End If
        
        
        versionName = getVersionNameById(versionID) & "-" & smallVersionID
        monthRange = getMonthRange(dateSeason)
        headInfo = nsrbm & "," & versionName & "," & dateYear & "," & monthRange
        allValueStr = headInfo & "," & allValueStr
        
        getExportData = allValueStr
'        Set fileObj = CreateObject("Scripting.FileSystemObject")
'        Set writeObj = fileObj.CreateTextFile(exportPath & "\" & fileName, True)
'        writeObj.WriteLine (headinfo & "," & allValueStr)
'        writeObj.Close
'        MsgBox "导出成功！"
        
        

    
    '        main_form.CB_Year.Text = dateYear
    '        main_form.CB_Season.Text = dateSeason
    '        Unload Me
    

End Function
Public Function getExportValuesById(id As String)
    Dim valueRs As ADODB.Recordset '保存报表输入项结果集
    Dim rownum As String
    Dim colName As String
    Dim value As String

    allValueStr = ""
    Call check_condatabase
    sql = "select row_num,col_num,value from t_baobiao_value where bb_content_id = '" & id & "' order by col_num,row_num"
    Set valueRs = cn.Execute(sql)
    
    While Not valueRs.EOF
                rownum = CStr(valueRs("row_num"))
            colnum = CStr(valueRs("col_num"))
            value = valueRs("value")
            
            allValueStr = allValueStr + value & ","  '对于为访问 Random 或 Binary 而打开的文件，直到最后一次执行的 Get 语句无法读出完整的记录时，EOF 都返回 False。
            If valueRs.EOF = False Then
            
            valueRs.MoveNext
            End If
    Wend
    
    
    valueRs.Close
    Set valueRs = Nothing
    'MsgBox allValueStr
    'allValueStr = Mid(allValueStr, 1, Len(allValueStr) - 1)
    'MsgBox allValueStr
    getExportValuesById = allValueStr
End Function

Public Function delBaobiaoByName(baobiaoName As String)
    Dim sql As String
    Call check_condatabase
    sql = "delete from t_baobiao where baobiao_name ='" & baobiaoName & "'"
    cn.Execute (sql)
End Function

'根据纳税人和所属期  获取 报表版本   因为一个纳税人一个所属期只能录入一个版本
Public Function get_version(nsrbm As String, dateYear As String, dateSeason As String) As String    'exportbj为1  则说明是导出或导入  只查询ID   为0则说明是保存  查询ID  有则返回 无则插入后生成ID
Dim flag As Boolean
Dim rs As ADODB.Recordset
Dim sql As String
Dim versionID As String
'Dim smallVersionID As String

sql = "select * from t_baobiao_content where nsrbm = '" & nsrbm & "' and user_name='" & username & "' and date_year = '" & dateYear & "' and date_season = '" & dateSeason & "'"
Call check_condatabase
Set rs = cn.Execute(sql)
If Not rs.EOF Then
  versionID = rs("version")
  'smallVersionID = rs("small_version_id")
End If
rs.Close
Set rs = Nothing

If versionID <> "" Then
    get_version = versionID
Else
   get_version = ""
End If


End Function


'根据版本号和小版本号  获取报表的名称
Public Function get_baobiao_name(verionID As String) As String     'exportbj为1  则说明是导出或导入  只查询ID   为0则说明是保存  查询ID  有则返回 无则插入后生成ID
Dim flag As Boolean
Dim rs As ADODB.Recordset
Dim sql As String
Dim baobiaoName As String

sql = "select * from t_baobiao where version_id = '" & verionID & "'"
Call check_condatabase
Set rs = cn.Execute(sql)
If Not rs.EOF Then
  baobiaoName = rs("baobiao_name")
End If
rs.Close
Set rs = Nothing

If baobiaoName <> "" Then
    get_baobiao_name = baobiaoName
Else
   get_baobiao_name = ""
End If


End Function
Public Function loadbaobiao1(bb_name As String)
    
     Dim baobiaoUrl As String
     Dim baobiaoName As String
     Dim aRs As ADODB.Recordset '保存纳税人的结果集
     Dim sql As String
     Dim hymxrs As ADODB.Recordset
     


     sql = "select baobiao_name from t_baobiao where version_id ='" & getVersionID(bb_name) & "' and baobiao_zl='" & hy & "'   "
      Set aRs = cn.Execute(sql)
      If Not aRs.EOF Then
        baobiaoName = aRs("baobiao_name")
        baobiaoUrl = App.Path & "\" & baobiaoName
        getUrl baobiaoUrl, main_form.F1Book1
       ' Me.F1Book1.ShowTabs = F1TabsOff
        'Me.F1Book1.AllowDesigner = False
        
         If hy <> "3" Then
  
  pid3 = main_form.F1Book1.ObjCreate(F1ObjDropDown, 2, 2, 3, 15)
  
  sql = " select cymc from t_cy where hy='" & hy & " ' order by id"
  
  
  Set hymxrs = cn.Execute(sql)
  
  
  While Not hymxrs.EOF
  
  main_form.F1Book1.ObjAddItem pid3, hymxrs("cymc")
  
  hymxrs.MoveNext
  
  Wend
  
 
  End If
        
          pid = main_form.F1Book1.ObjCreate(F1ObjDropDown, 2, 8, 3, 15)
           main_form.F1Book1.ObjAddItem pid, "扩大经营"
           main_form.F1Book1.ObjAddItem pid, "正常经营"
           main_form.F1Book1.ObjAddItem pid, "减产经营"
           main_form.F1Book1.ObjAddItem pid, "半停产"
           main_form.F1Book1.ObjAddItem pid, "停产"
           'Me.F1Book1.ObjItem(pid, 0)
           'Me.F1Book1.ObjValue (pid)
           pid2 = main_form.F1Book1.ObjCreate(F1ObjDropDown, 2, 9, 3, 15)
           main_form.F1Book1.ObjAddItem pid2, "国税"
           main_form.F1Book1.ObjAddItem pid2, "地税"
           main_form.F1Book1.ObjAddItem pid2, "国税地税混合"
           main_form.F1Book1.ObjAddItem pid2, "不征收"
           main_form.F1Book1.MaxRow = 70
           
           Dim i As Integer
           For i = 1 To main_form.F1Book1.NumSheets
           main_form.F1Book1.Sheet = i
           main_form.F1Book1.ShowLockedCellsError = False
              main_form.F1Book1.BackColor = &HFFFF00
            main_form.F1Book1.MaxRow = 70
            main_form.F1Book1.AllowEditHeaders = False
           Next i
           main_form.F1Book1.Sheet = 1
          
        End If
End Function

Public Function change_col(Col As String) As String '列转换
Dim C_ol As String
If Col = "3" Then
   C_ol = "C"
ElseIf Col = "4" Then
   C_ol = "D"
ElseIf Col = "7" Then
   C_ol = "G"
ElseIf Col = "8" Then
   C_ol = "H"
End If
change_col = C_ol




End Function

Public Function change_coll(a As String) As Long
If a = "A" Then
change_coll = 1
ElseIf a = "B" Then
change_coll = 2
ElseIf a = "C" Then
change_coll = 3
ElseIf a = "D" Then
change_coll = 4
ElseIf a = "E" Then
change_coll = 5
ElseIf a = "F" Then
change_coll = 6
ElseIf a = "G" Then
change_coll = 7
ElseIf a = "H" Then
change_coll = 8
End If

End Function

