Attribute VB_Name = "ValidateModel"


'������֤������֤���ݺϷ���
Public Function validate_exp_data(dateYear As String, dateSeason As String, nsrbm As String, textWrongMes As TextBox, ScriptControl1 As ScriptControl, F1Book1 As F1Book) As Boolean

textWrongMes.Text = ""
textWrongMes.Text = textWrongMes.Text & pd_hanzi(dateYear, dateSeason, nsrbm)
If Not textWrongMes.Text = "" Then
validate_exp_date = False
Else



Dim dataRange As String   '�����ʾ������ķ�Χ������  ��1��1��3��3 �������A1��C3�������ÿ����Ԫ���ŵĶ���һ�����뷶Χ���ݴ�ÿ�������sheet2��
Dim valueRange As String   '���汨���������ķ�Χ
dataRange = getData4IndexSheet(1, 5500, 1, 5500, 1, F1Book1) '�̶��ڶ��е�һ�б���������������������Ϣ
'ע��˴��Ľش�  ���ڲ�ͬ��Χ��صĳ��Ȳ�ͬ ���������ݵķ�Χ�й�  �����Ǵ�5100��1�п�ʼ��ţ������������£�����ֻҪ����5999��  �˹��򼴿���
dataRange = Mid(dataRange, 8, Len(dataRange) - 8)   '  ��dataRange = 5100,1,12,1,18,1  ����5100��1�е�ֵΪ12��1��18��1  ��������ش����ǻ��12,1,18,1
'MsgBox sheetName & dataRange
Dim rowColArray
rowColArray = Split(dataRange, ",")

If UBound(rowColArray) - LBound(rowColArray) + 1 = 4 Then   '����������ַ������飬��ʽΪstartrow��startcol��endrow��endcol  ���Գ��ȱ���Ϊ4

    Dim rowIndex As Integer
    Dim rowInfoArrayLen As Integer
    Dim rowInfo As String
    
    Dim valueInfoArray
    Dim firstFlag As String
    
    Dim realValue As Double   '����Ҫ�Ƚϵ�ֵ
    Dim mesStr As String
    'ȡֵʱ�Ĳ���
    Dim sheetName As String
    Dim paramYear As String
    Dim paramSeason As String
    Dim paramRow As String
    Dim paramCol As String
    
    Dim czf As String
    
    
    
    Dim startRow As Integer
    Dim endRow As Integer
    Dim rowStr As String
    Dim maxColNum As Integer
    
    Dim rowInfoArray  '  ' 5501,1,0:�ֽ�������:0:0:40:3;5501,2,1:1;5501,3,0:�ֽ�������:0:0:41:4  ���������ɵ�����
    startRow = CInt(rowColArray(0))
    endRow = CInt(rowColArray(2))
    maxColNum = CInt(rowColArray(3))
    'MsgBox dataRange
    ' Exit Sub
    If Trim(dataRange) <> "" Then
      For rowIndex = startRow To endRow     'ѭ��ÿһ��
        rowStr = getData4IndexSheet(1, rowIndex, 1, rowIndex, maxColNum, F1Book1)
        ' 5501,1,0:�ֽ�������:0:0:40:3;5501,2,1:1;5501,3,0:�ֽ�������:0:0:41:4
        rowInfoArray = Split(rowStr, ";")
        Dim rowRealStr As String
        rowRealStr = ""
        For rowInfoArrayLen = LBound(rowInfoArray) To UBound(rowInfoArray)   'ѭ��ÿ�е�ÿ��
    
    
            rowInfo = rowInfoArray(rowInfoArrayLen)
            If rowInfo <> "" Then
                rowInfo = Mid(rowInfo, 8, Len(rowInfo) - 7)
                If rowInfo <> "" Then
                    'MsgBox "ִ�е�һ�У�" & rowInfo
                    valueInfoArray = Split(rowInfo, ":")
               
                    firstFlag = valueInfoArray(0)
                    'MsgBox "firstFlag:" & firstFlag
                    
                    If firstFlag = "0" Then
                        sheetName = valueInfoArray(1)
                        paramYear = valueInfoArray(2)
                        If paramYear = "2" Then '�жϸ������Ƿ���ȡ����ֵ
                        paramYear = "0"
                        jdz = 1 'ȡ����ֵ���
                        Else
                        jdz = 0 '��ȡ����ֵ
                        End If
                        If paramYear = "3" Then
                        paramYear = "0"
                        fh = 1
                        Else
                        fh = 0
                        End If
                        paramSeason = valueInfoArray(3)
                        paramRow = valueInfoArray(4)
                        paramCol = valueInfoArray(5)
                        realValue = getValueByParam(sheetName, paramYear, paramSeason, paramRow, paramCol, nsrbm, dateYear, dateSeason) '���ݵ�Ԫ���еĲ�����ȡֵ
                        'MsgBox "realValue:" & realValue
                  
                        
                        
                              If realValue = 1.0000001 Then
                        GoTo ok
                       End If
                        rowRealStr = rowRealStr & realValue
                        If sn = 1 Then
                        sheetName = "����� " & sheetName
                        ElseIf sjd = 1 Then
                        sheetName = "�ϼ���" & sheetName
                        End If
                        
                        
                        
                        
                        If jdz = 1 Then
                        
                        mesStr = mesStr & sheetName & "�� " & paramRow & "��" & change_col(paramCol) & "�еľ���ֵ��" & realValue
                        Else
                         mesStr = mesStr & sheetName & "�� " & paramRow & "��" & change_col(paramCol) & "�е�ֵ��" & realValue
                         End If
                   '���ݵ�Ԫ��������  ��ȡ������   �����䣨���ǽ������������
                    ElseIf firstFlag = "1" Then
                        czf = valueInfoArray(1)
                        If czf = "1" Then
                            czf = "="
                        ElseIf czf = "2" Then
                            czf = ">"
                        ElseIf czf = "3" Then
                            czf = "<"
                        ElseIf czf = "4" Then
                            czf = ">="
                        ElseIf czf = "5" Then
                            czf = "<="
                        ElseIf czf = "6" Then
                            czf = "+"
                        ElseIf czf = "7" Then
                            czf = "-"
                        End If
                        rowRealStr = rowRealStr & czf
                        mesStr = mesStr & czf
                    End If
                    
                Else
                End If
                
                
            End If
            
          
          
        Next
       
   ' MsgBox "���" & rowRealStr
    If rowRealStr <> "" Then
    
        Dim ifStr As String      'msgbox ""
        'ifStr = "if " & rowRealStr & " then" & vbCrLf & "MsgBox ""True"" " & vbCrLf & "  Else " & vbCrLf & " MsgBox ""False "" & """ & mesStr & """" & vbCrLf & "End If"
        'ƴ���߼��ַ���  ��ӡ������Ϣ
        ifStr = "if " & rowRealStr & " then" & vbCrLf & "  Else " & vbCrLf & "  textWrongMes.Text = textWrongMes.Text & "" " & nsrbm & " " & mesStr & """ & vbCrlf " & vbCrLf & "End If"
        'MsgBox ifStr
        ScriptControl1.ExecuteStatement (ifStr)  'ִ��ƴ�����ַ����߼����ʽ
         
    End If
ok:
    
    mesStr = ""
    'Exit For
    

    Next
   
    
    
    Else
    MsgBox "�˱���û�������Χ��Ϣ���޷�����"
    End If

Else
    validate_exp_data = False
End If


If textWrongMes.Text = "" Then
    validate_exp_data = True
    textWrongMes.Text = "���������Ѿ�ͨ��У�飡�ɵ���"
Else
   validate_exp_data = False
   
End If
End If

End Function

'�����ǽ�����ʾ�����

Public Function validate_ts_data(dateYear As String, dateSeason As String, nsrbm As String, textWrongMes As TextBox, ScriptControl2 As ScriptControl, F1Book1 As F1Book) As Boolean
If pd_hanzi(dateYear, dateSeason, nsrbm) <> "" Then

Else

'textWrongMes.Text = ""
Dim dataRange As String   '�����ʾ������ķ�Χ������  ��1��1��3��3 �������A1��C3�������ÿ����Ԫ���ŵĶ���һ�����뷶Χ���ݴ�ÿ�������sheet2��
Dim valueRange As String
dataRange = getData4IndexSheet(1, 5600, 1, 5600, 1, F1Book1)
dataRange = Mid(dataRange, 8, Len(dataRange) - 8)
Dim rowColArray
rowColArray = Split(dataRange, ",")

If UBound(rowColArray) - LBound(rowColArray) + 1 = 4 Then   '����������ַ������飬��ʽΪstartrow��startcol��endrow��endcol  ���Գ��ȱ���Ϊ4

    Dim rowIndex As Integer
    Dim rowInfoArrayLen As Integer
    Dim rowInfo As String
    
    Dim valueInfoArray
    Dim firstFlag As String
    
    Dim realValue As Double   '����Ҫ�Ƚϵ�ֵ
    Dim mesStr As String
    'ȡֵʱ�Ĳ���
    Dim sheetName As String
    Dim paramYear As String
    Dim paramSeason As String
    Dim paramRow As String
    Dim paramCol As String
    
    Dim czf As String
    
    
    
    Dim startRow As Integer
    Dim endRow As Integer
    Dim rowStr As String
    Dim maxColNum As Integer
    Dim rowInfoArray
    
    '  ' 5501,1,0:�ֽ�������:0:0:40:3;5501,2,1:1;5501,3,0:�ֽ�������:0:0:41:4  ���������ɵ�����
    startRow = CInt(rowColArray(0))
    endRow = CInt(rowColArray(2))
    maxColNum = CInt(rowColArray(3))
    'MsgBox dataRange
    ' Exit Sub
    If Trim(dataRange) <> "" Then
      For rowIndex = startRow To endRow     'ѭ��ÿһ��
        rowStr = getData4IndexSheet(1, rowIndex, 1, rowIndex, maxColNum, F1Book1)
        ' 5501,1,0:�ֽ�������:0:0:40:3;5501,2,1:1;5501,3,0:�ֽ�������:0:0:41:4
        rowInfoArray = Split(rowStr, ";")
        Dim rowRealStr As String
        rowRealStr = ""
        For rowInfoArrayLen = LBound(rowInfoArray) To UBound(rowInfoArray)   'ѭ��ÿ�е�ÿ��
    
    
            rowInfo = rowInfoArray(rowInfoArrayLen)
            If rowInfo <> "" Then
                rowInfo = Mid(rowInfo, 8, Len(rowInfo) - 7)
                If rowInfo <> "" Then
                    'MsgBox "ִ�е�һ�У�" & rowInfo
                    valueInfoArray = Split(rowInfo, ":")
               
                    firstFlag = valueInfoArray(0)
                    'MsgBox "firstFlag:" & firstFlag
                    
                    If firstFlag = "0" Then
                        sheetName = valueInfoArray(1)
                        paramYear = valueInfoArray(2)
                        If paramYear = "3" Then
                        paramYear = "0"
                        fh = 1
                        Else
                        fh = 0
                        End If
                        
                        paramSeason = valueInfoArray(3)
                        paramRow = valueInfoArray(4)
                        paramCol = valueInfoArray(5)
                        realValue = getValueByParam(sheetName, paramYear, paramSeason, paramRow, paramCol, nsrbm, dateYear, dateSeason) '���ݵ�Ԫ���еĲ�����ȡֵ
                        'MsgBox "realValue:" & realValue
                       If realValue = 1.0000001 Then
                        GoTo o_k
                       End If
                        rowRealStr = rowRealStr & realValue
                        If sn = 1 Then
                        sheetName = "����� " & sheetName
                        ElseIf sjd = 1 Then
                        sheetName = "�ϼ���" & sheetName
                        End If
                        
                        
                        
                        If fh = 1 Then
                        
                        mesStr = mesStr & sheetName & "�� " & paramRow & "��" & change_col(paramCol) & "�е�ֵ���෴����" & realValue
                        Else
                         mesStr = mesStr & sheetName & "�� " & paramRow & "��" & change_col(paramCol) & "�е�ֵ��" & realValue
                         End If
                   '���ݵ�Ԫ��������  ��ȡ������   �����䣨���ǽ������������
                    ElseIf firstFlag = "1" Then
                        czf = valueInfoArray(1)
                        If czf = "1" Then
                            czf = "="
                        ElseIf czf = "2" Then
                            czf = ">"
                        ElseIf czf = "3" Then
                            czf = "<"
                        ElseIf czf = "4" Then
                            czf = ">="
                        ElseIf czf = "5" Then
                            czf = "<="
                        ElseIf czf = "6" Then
                            czf = "+"
                        ElseIf czf = "7" Then
                            czf = "-"
                        End If
                        rowRealStr = rowRealStr & czf
                        mesStr = mesStr & czf
                    End If
                    
                Else
                End If
                
                
            End If
            
          
          
        Next
       
   ' MsgBox "���" & rowRealStr
  
    If rowRealStr <> "" Then
     Dim ifStr As String
           'msgbox ""
        'ifStr = "if " & rowRealStr & " then" & vbCrLf & "MsgBox ""True"" " & vbCrLf & "  Else " & vbCrLf & " MsgBox ""False "" & """ & mesStr & """" & vbCrLf & "End If"
        'ƴ���߼��ַ���  ��ӡ������Ϣ
        ifStr = "if " & rowRealStr & " then" & vbCrLf & "  Else " & vbCrLf & "  textWrongMes.Text = textWrongMes.Text & "" " & nsrbm & " " & mesStr & """ & vbCrlf" & vbCrLf & "End If"
        'MsgBox ifStr
      
         ScriptControl2.ExecuteStatement (ifStr)
    
    End If
o_k:
 mesStr = ""
    'Exit For

    Next
End If
Else
    validate_ts_data = False
End If

If textWrongMes.Text = "" Then
    validate_ts_data = True
    'textWrongMes.Text = "���������Ѿ�ͨ����ʾ��У�飡"
Else
   validate_ts_data = False
   
End If

 Dim mess As String
 
mess = pdzero(nsrbm, dateYear, dateSeason) '�жϱ�������Ƿ������ (�¾��ʲ���ծ��)
Form_Export.text_warning_mes = Form_Export.text_warning_mes & mess
mess = pd_price(nsrbm, dateYear, dateSeason) '�ж�
Form_Export.text_warning_mes = Form_Export.text_warning_mes & mess
mess = ""
End If
End Function
Public Function getValueByParam(ByVal sheetName As String, ByVal paramYear As String, ByVal paramSeason As String, ByVal paramRow As String, ByVal paramCol As String, ByVal nsrbm As String, ByVal dateYear As String, ByVal dateSeason As String) As Double
Dim realValue As Integer
realValue = 0

'�������   0����   1����
If paramYear = "0" Then
sn = 0
Else
dateYear = CStr(CInt(dateYear) - 1)
sn = 1
End If

'���paramSeasonΪ0  ���ʾ�Ǳ�����  ��Ϊ0  ���ʾ����ָ���� �� 1��2��3��4  ����һ��
If paramSeason = "0" Then
sjd = 0
ElseIf paramSeason = "5" Then

dateSeason = CStr(CInt(dateSeason) - 1)
sjd = 1
Else
dateSeason = paramSeason
End If


Dim bb_content_id As String

bb_content_id = getBbContentID(nsrbm, sheetName, dateYear, dateSeason)

Call check_condatabase
sql = "select value from t_baobiao_value where bb_content_id = '" & bb_content_id & "' and row_num = " & CInt(paramRow) & " and col_num = " & CInt(paramCol)
Set valueRs = cn.Execute(sql)
    
 If Not valueRs.EOF Then
    'MsgBox valueRs("value")
    If valueRs("value") = "" Then
        getValueByParam = 1.0000001
    Else
        getValueByParam = CDbl(valueRs("value"))
    End If
    If jdz = 1 Then
    getValueByParam = Abs(getValueByParam)
    Else
    End If
    If fh = 1 And getValueByParam <> 1.0000001 Then
    getValueByParam = -(getValueByParam)
    Else
    End If
 Else
    getValueByParam = 1.0000001
 End If
  
  valueRs.Close
  Set valueRs = Nothing




End Function

Public Function getCzfByFlag(czfStr As String) As Integer
Dim czf As String
inputValue = 0
getValueByRange = inputValue



End Function

Public Function getBbContentID(nsrbm As String, sheetName As String, dateYear As String, dateSeason As String) As String
    Dim rs As ADODB.Recordset  '������˰�˵Ľ����
    Dim sql As String
    Dim baobiaoID As String
    
    Call check_condatabase
    sql = "select id from t_baobiao_content where user_name='" & username & "' and nsrbm = '" & nsrbm & "' and baobiao_name = '" & sheetName & "' and date_year = '" & dateYear & "' and date_season = '" & dateSeason & "'"
    'MsgBox sql
    Set rs = cn.Execute(sql)
    If Not rs.EOF Then
        baobiaoID = rs("id")
        getBbContentID = baobiaoID
    Else
    getBbContentID = ""
    'MsgBox "���ر�������������ݿ�"
    End If
   
End Function
Public Function pdzero(nsrbm As String, date_year As String, date_season As String) '�ж��ǲ��Ǵ���0
Dim baobiaoID As String
baobiaoID = getBbContentID(nsrbm, "�ʲ���ծ��", date_year, date_season)
 Dim valueRs As ADODB.Recordset '���汨������������
    Dim rownum As String
    Dim colName As String
    Dim value As Double
    Dim mess As String
   mess = ""
    Call check_condatabase
    sql = "select row_num,col_num,value from t_baobiao_value where bb_content_id = '" & baobiaoID & "'"
    Set valueRs = cn.Execute(sql)
    
    While Not valueRs.EOF
        rownum = CStr(valueRs("row_num"))
        colnum = CStr(valueRs("col_num"))
        If valueRs("value") = "" Then
        value = 0
        Else
       
        value = CDbl(valueRs("value"))
        End If
      
        
        
        If value < 0 Then
        mess = mess & " " & nsrbm & " " & "�ʲ���ծ���� " & rownum & "��" & change_col(CStr(colnum)) & "�е�ֵ��" & value & "<0" & vbCrLf
       Else
       End If
         valueRs.MoveNext
    Wend
    valueRs.Close
    pdzero = mess
End Function

Public Function pd_price(nsrbm As String, date_year As String, date_season As String) '�жϼ۸�
Dim baobiaoID As String
Dim baobiaoID2 As String
Dim mes As String
If date_year = "2009" And date_season = "1" Then
pd_price = ""
ElseIf date_year <> "2009" And date_season = "1" Then
baobiaoID = getBbContentID(nsrbm, "��Ӫ��Ϣ��", date_year, date_season)
date_year = CStr(CInt(date_year) - 1)
date_season = "4"
baobiaoID2 = getBbContentID(nsrbm, "��Ӫ��Ϣ��", date_year, date_season)
mes = pd_cg(baobiaoID, baobiaoID2)
If mes = "" Then
pd_price = ""
Else
pd_price = " " & nsrbm & " " & pd_cg(baobiaoID, baobiaoID2)
End If
Else
baobiaoID = getBbContentID(nsrbm, "��Ӫ��Ϣ��", date_year, date_season)
date_season = CStr(CInt(date_season) - 1)
baobiaoID2 = getBbContentID(nsrbm, "��Ӫ��Ϣ��", date_year, date_season)
mes = pd_cg(baobiaoID, baobiaoID2)
If mes = "" Then
pd_price = ""
Else
pd_price = " " & nsrbm & " " & pd_cg(baobiaoID, baobiaoID2)
 End If
 End If
End Function

Public Function pd_cg(a As String, B As String) '�жϼ۸񸡶��ǲ��ǳ���50%
 Dim valueRs As ADODB.Recordset
 Dim value2Rs As ADODB.Recordset
Dim value1 As Double
Dim value2 As Double


Call check_condatabase
    sql = "select value from t_baobiao_value where bb_content_id = '" & a & "'and row_num=5"
    Set valueRs = cn.Execute(sql)
    If Not valueRs.EOF Then
    If valueRs("value") = "" Then
        value1 = 1.000001
        Else
       
        value1 = CDbl(valueRs("value"))
        End If
    Else
    
    End If
        valueRs.Close
  Call check_condatabase
    sql = "select value from t_baobiao_value where bb_content_id = '" & B & "'and row_num=5"
    Set value2Rs = cn.Execute(sql)
    If Not value2Rs.EOF Then
    If value2Rs("value") = "" Then
        value2 = 1.000001
        Else
       
        value2 = CDbl(value2Rs("value"))
        End If
    
        value2Rs.Close
        If value1 <> 1.000001 And value2 <> 1.000001 Then
        
        If Abs(value1 - value2) > 0.5 * value2 Then
        pd_cg = "��Ӫ��Ϣ���� " & "5��C�е�ֵ��" & value1 & "������" & value2 & "�仯����50%"
        Else
        End If
        Else
        End If
    Else
    End If
End Function
Public Function pd_hanzi(dateYear As String, dateSeason As String, nsr As String) As String '�ж��ǲ��������˲������ֵ�ֵ,hzbj=0��ʾû��,hzbj=1��ʾ�в��Ϸ������롣
Dim v_Rs As ADODB.Recordset
Dim sql1 As String
Dim sql2 As String
Dim l_Rs As ADODB.Recordset
Dim value As String
Dim r As String
Dim c As String
Dim hzbj As Integer
hzbj = 0

Call check_condatabase
sql1 = "select id from t_baobiao_content where nsrbm='" & nsr & " 'and date_year='" & dateYear & "' and date_Season='" & dateSeason & "'"
Set v_Rs = cn.Execute(sql1)
While Not v_Rs.EOF
bb_id = v_Rs("id")
sql2 = "select * from t_baobiao_value where bb_content_id='" & bb_id & "'"
Set l_Rs = cn.Execute(sql2)
While Not l_Rs.EOF
If l_Rs("value") = "" Then
value = 0
Else
value = l_Rs("value")
End If

r = l_Rs("row_num")
c = l_Rs("col_num")
If Not IsNumeric(value) Then
MsgBox (bb_id)

MsgBox (r)
MsgBox (c)
hzbj = 1
Else
End If

l_Rs.MoveNext
Wend
v_Rs.MoveNext
Wend
If hzbj = 1 Then
Dim mss As String
 



pd_hanzi = "���ݲ��Ϸ�,�����Ƿ������˲�����ֵ�����ݣ����磺���֣��ַ��ȣ�"
Else
pd_hanzi = ""
End If
End Function




