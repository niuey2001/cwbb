Attribute VB_Name = "global"
Public cn As ADODB.Connection 'ȫ�ֵ����ݿ����Ӷ���
Public czry_flag As String '�����û�Ȩ���ַ���
Public isxg As Boolean '������û������޸��û�
Public nodename As String '����ĳ���û����ʻ���
Public rspsw As ADODB.Recordset '�����û��ʺ��������õĽ������
Public database_data As String '���浼�����ݱ���Ϣʱ��access���ݿ�����
Public username As String '�����½���û���
Public userType As String
Public pid, pid2, pid3 As Long 'formula one �пؼ����ر��
Public xx, yy, zz As String
Public jz_bj As Integer '�����������ݵ�ʱ���Ƿ��Ǿɰ汾�������
Public sn, sjd As Integer '��ʾ��ʱ���Ƿ���ʾ�����
Public jdz As Integer '�Ƿ�ȡ����ֵ
Public fh As Integer '�Ƿ�ȡ��
Public hy  As String '�ж���ҵ
 









Public Sub condatabase()           '�������ӵ�feiyong���ݿ�ļ�¼Դ                                     ���ӱ������ݿ�JIMMY
    Set cn = New ADODB.Connection
      '  cn.Provider = "sqloledb"
      '  cn.Properties("Data Source").Value = "JIMMY"       �����뱾�����ݿ������
      '  cn.Properties("Initial Catalog").Value = "YAOFEI"   ���ݿ������
      '  cn.Properties("Integrated Security").Value = "SSPI"
      
      cn.ConnectionString = "provider=Microsoft.Jet.OLEDB.4.0;Data source =" & App.Path & "\financialForm.mdb" '��ת����access���ݿ�·��
      cn.Properties("Jet OLEDB:Database Password") = "niuey" 'ACCESS ����
      cn.Open
End Sub

Public Sub check_condatabase()
If cn.State = 1 Then 'cn.State��ֵΪ1��ʾ���ݿ⴦������״̬
Else
   Call condatabase
End If
End Sub

Public Function date_change(sea_son As String) As String '�±伾��
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

Public Function change_date(sea_son As String) As String '���ڱ���
Dim sql As String
Dim rs As ADODB.Recordset
sql = "select * from t_month_dm where season_id = '" & sea_son & " '"
Call check_condatabase
Set rs = cn.Execute(sql)
change_date = rs("month_range")
rs.Close
End Function


Public Sub close_condatabase()   '�ر�����Դ
If cn.State = 1 Then
   cn.Close
End If
End Sub
'�����ļ�·��   ��ȡtxt�ļ�������
Public Function getTxt(txtPath As String) As String
Dim i As Integer: i = FreeFile   '����һ�������ļ���
Open txtPath For Input As #i
    getTxt = StrConv(InputB(LOF(i), i), vbUnicode)  ' ����ϵͳ��ȱʡ��ҳ���ַ���ת��  ��Ӣ�Ķ���
Close #i
End Function
'�����ļ�·�����к�   ��ȡtxt�ļ�������һ��
Public Function getLine(txtPath As String, lineNum As Integer) As String
getLine = Split(getTxt(txtPath), vbCrLf)(lineNum - 1)
End Function

'fileNum = FreeFile '����һ������ļ���
'Open App.Path & "\user_info.ini" For Input As #fileNum  '���ļ�"
'Open file_path & "\��˰����Ϣ.txt" For Input As #fileNum  '���ļ�"
'Do While Not EOF(fileNum)
'Line Input #fileNum, text_line
'array_user_info = Split(text_line, ",")
'MsgBox text_line
'Loop
'Close #fileNum   '�ر��ļ�
'�����ļ�·�����к�   ��ȡtxt�ļ�������һ��

'��ȡ�ļ�������  ����ɵ�����
Public Function getLineArray(txtPath As String) As String()
Dim linearray() As String
linearray = Split(getTxt(txtPath), vbCrLf)  '���س������ַ�������
getLineArray = linearray
End Function
'�� 2��4��sfsdf;  ���ݴ��ַ������sfsdf
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

'ͨ��version�����
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
'ͨ��version�����
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
    Dim rs As ADODB.Recordset '������˰�˵Ľ����
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
    'MsgBox "���ر�������������ݿ�"
    End If
   
End Function

Public Function getValuesById(id As String)
    Dim valueRs As ADODB.Recordset '���汨������������
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
    qybj = "1"   'Ĭ��Ϊ1  ������
    
        Dim versionID As String
        versionID = getVersionID(versionName)
        baobiaoID = getBaobiaoID_2(versionID, hyzl)
        If baobiaoID <> "" Then
              
             '���ݿ����
                Dim sql As String
                Dim versionRs As ADODB.Recordset
                
                Call check_condatabase
                sql = "select nsrbm from t_baobiao_version where nsrbm = '" & nsrbm & "' and user_name='" & username & "'"
                
                Set versionRs = cn.Execute(sql)
                If Not versionRs.EOF Then
                    Dim choose As Integer
                    choose = MsgBox("����汾�ѱ����ã��Ƿ����", vbOKCancel)
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
        MsgBox "�˰汾������Ϣ�����ڣ�"
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
              
            valueRs.AddNew '��ӱ�����Ϣ
            valueRs("user_name") = username
            valueRs("nsrbm") = nsrbm
            valueRs("baobiao_id") = baobiaoID
            valueRs("qybj") = qybj
            valueRs.Update
            valueRs.Close
            Set valueRs = Nothing
           ' MsgBox "����ɹ���"
End Sub
 
 

 


Public Sub updatensr_hy(hy As String, bm As String)

 Dim hyRs As ADODB.Recordset
Dim strsql As String
Dim hy_xg As String

 If hy = "��ҵ��ҵ" Then
 hy_xg = "1"
 ElseIf hy = "���ز���ҵ" Then
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
  
    hyRs.Fields("zchy") = hy_xg '�Ƿ������� ����
        hyRs.Update
       
       Else
       
        End If
End Sub

Public Function getExportData(nsrbm As String, dateStr As String) As String
  Dim dateYear As String
   Dim dateSeason As String


    Dim bbValueStr As String  'һ�����������ֵ
    
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
'        MsgBox "�����ɹ���"
        
        

    
    '        main_form.CB_Year.Text = dateYear
    '        main_form.CB_Season.Text = dateSeason
    '        Unload Me
    

End Function
Public Function getExportValuesById(id As String)
    Dim valueRs As ADODB.Recordset '���汨������������
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
            
            allValueStr = allValueStr + value & ","  '����Ϊ���� Random �� Binary ���򿪵��ļ���ֱ�����һ��ִ�е� Get ����޷����������ļ�¼ʱ��EOF ������ False��
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

'������˰�˺�������  ��ȡ ����汾   ��Ϊһ����˰��һ��������ֻ��¼��һ���汾
Public Function get_version(nsrbm As String, dateYear As String, dateSeason As String) As String    'exportbjΪ1  ��˵���ǵ�������  ֻ��ѯID   Ϊ0��˵���Ǳ���  ��ѯID  ���򷵻� ������������ID
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


'���ݰ汾�ź�С�汾��  ��ȡ���������
Public Function get_baobiao_name(verionID As String) As String     'exportbjΪ1  ��˵���ǵ�������  ֻ��ѯID   Ϊ0��˵���Ǳ���  ��ѯID  ���򷵻� ������������ID
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
     Dim aRs As ADODB.Recordset '������˰�˵Ľ����
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
           main_form.F1Book1.ObjAddItem pid, "����Ӫ"
           main_form.F1Book1.ObjAddItem pid, "������Ӫ"
           main_form.F1Book1.ObjAddItem pid, "������Ӫ"
           main_form.F1Book1.ObjAddItem pid, "��ͣ��"
           main_form.F1Book1.ObjAddItem pid, "ͣ��"
           'Me.F1Book1.ObjItem(pid, 0)
           'Me.F1Book1.ObjValue (pid)
           pid2 = main_form.F1Book1.ObjCreate(F1ObjDropDown, 2, 9, 3, 15)
           main_form.F1Book1.ObjAddItem pid2, "��˰"
           main_form.F1Book1.ObjAddItem pid2, "��˰"
           main_form.F1Book1.ObjAddItem pid2, "��˰��˰���"
           main_form.F1Book1.ObjAddItem pid2, "������"
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

Public Function change_col(Col As String) As String '��ת��
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

