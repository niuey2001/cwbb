Attribute VB_Name = "fomularone"
'ȫ�ֱ���
Public isValidateFailed As Boolean       '��֤�Ƿ�ͨ���ı�־λ��true��Ϊͨ����false��ͨ��

'Public isShowData As Boolean
Public xlApp As New excel.Application

Public validateFailedMes As String       '��֤ʧ��ʱ�洢�Ĵ�����ʾ��Ϣ������vts�ж������Ϣ

Public isGetErrorMes As Boolean          '��֤ʧ���¼��жϵ�ǰ������֤����Դ�Ƿ�ΪgetErrorMes����

Public validateFailedColor               '��֤Ϊͨ��ʱ��cell��ʾ����ɫ��Ĭ��Ϊ��ɫ
'ȡ��excelģ��
Public Function getUrl(theUrl As String, F1Book1 As F1Book)

If Dir(theUrl) = "" Then '�ļ�������
    MsgBox "�������ڣ�"
Else

 F1Book1.URL = theUrl
 
  
End If
   
End Function
'ȡ�ñ��ϵ����ݺ������꣬����Ϊstring����,�˷���Ĭ�ϴӵ�һ��sheetҳ��ȡ�����ݣ��ʺ��ڵ�sheet��excel
'exlQsdX,exlQsdYΪ���ݵ���ʼ������
'exlZdX,exlZdYΪ���ݵ��յ�����
Public Function getData(exlQsdX As Integer, exlQsdY As Integer, exlZdX As Integer, exlZdY As Integer, F1Book1 As F1Book) As String
    
    getData = getData4IndexSheet(1, exlQsdX, exlQsdY, exlZdX, exlZdY, F1Book1)

End Function

Public Function showData(showDataStr As String, F1Book1 As F1Book)
showData4IndexSheet 1, showDataStr, F1Book1
End Function

'ȡ�ñ��ϵ����ݺ������꣬����Ϊstring����,�˷���ָ����sheetҳ��ȡ�����ݣ��ʺ��ڶ�sheetҳ��excel
'sheetIndexΪҪȡ�����ݵ�sheet����ţ������Ϊ1��ʼ����
'exlQsdX,exlQsdYΪ���ݵ���ʼ������
'exlZdX,exlZdYΪ���ݵ��յ�����
Public Function getData4IndexSheet(sheetIndex As Integer, exlQsdX As Integer, exlQsdY As Integer, exlZdX As Integer, exlZdY As Integer, F1Book1 As F1Book) As String
    
    Dim copyArray() As Variant
    
    ReDim copyArray(exlQsdX To exlZdX, exlQsdY To exlZdY) As Variant
    
    F1Book1.CopyDataToArray sheetIndex, exlQsdX, exlQsdY, exlZdX, exlZdY, True, copyArray
    
   'Action = MsgBox(copyArray(9, 3), vbOKCancel, "ok")
   
   Dim returnString As String
   returnString = ""
   
   For X = exlQsdX To exlZdX
   
        For Y = exlQsdY To exlZdY
        
          '  If Trim(copyArray(X, Y)) <> "" Then
                 copyArray(X, Y) = Replace(copyArray(X, Y), " ", "")
                 returnString = returnString & X & "," & Y & "," & Trim(copyArray(X, Y)) & ";"
            
          '  End If
               
        Next Y
   
   Next X
        
   getData4IndexSheet = returnString

End Function
'չ������,�˷�������Ϊָ����ŵ�sheetҳչ�����ݣ��ʺ��ڶ�sheetҳ��excel
'sheetIndexΪҪչ�����ݵ�sheet����ţ������Ϊ1��ʼ����


Public Function showData4IndexSheet(sheetIndex As Integer, showDataStr As String, F1Book1 As F1Book)
'isShowData = True

'1 ��һ�ν����ַ�����array1()��������Ԫ��Ϊx,y,value �����ֵ꣬��
Dim array1() As String
array1() = Split(showDataStr, ";")


'2 �������鴦��ÿ�� ���ֵ꣬
Dim array2() As String
For i = 0 To UBound(array1) - 1
    '3 ��2�ν����ַ���,����ÿһ��ѭ����array2[0]Ϊ�����꣬array2[1]Ϊ�����꣬array2[2]Ϊֵ
    array2() = Split(array1(i), ",")
    
    
 '�жϴ˵�Ԫ���й�ʽ  �Ͳ���ֵ   lwy���
 If haveNoFormula(sheetIndex, CInt(array2(0)), CInt(array2(1)), F1Book1) Then
        
        '-------------------------
    'Dim cellformat As F1CellFormat
    'Set cellformat = F1Book1.CreateNewCellFormat
    'With cellformat
    '    .FontColor = vbBlack
    'End With
    
    F1Book1.Sheet = sheetIndex
    
    F1Book1.SetActiveCell array2(0), array2(1)
    
    'F1Book1.SetCellFormat cellformat
        
        '-------------------------
           
        '4 Ϊ����ֵ
    If array2(2) <> "" Then
    F1Book1.EntrySRC(sheetIndex, array2(0), array2(1)) = array2(2)
    'F1Book1.entr
    End If
    

End If



Next i

'isShowData = False
End Function


'ȡ�ñ��ϵ����ݺ������꣬����Ϊstring����,�˷���ָ����sheetҳ��ȡ�����ݣ��ʺ��ڶ�sheetҳ��excel
'sheetNameΪҪȡ�����ݵ�sheetҳ������
'exlQsdX,exlQsdYΪ���ݵ���ʼ������
'exlZdX,exlZdYΪ���ݵ��յ�����
Public Function getData4NameSheet(sheetName As String, exlQsdX As Integer, exlQsdY As Integer, exlZdX As Integer, exlZdY As Integer, F1Book1 As F1Book) As String
    Dim returnStr As String
    Dim sheetIndex As Integer
    
    '��������sheetNameװ���ɶ�Ӧ��sheetIndex
    sheetIndex = transName2Index(sheetName, F1Book1)
    
    '�������Ϊ0˵��û�ж�Ӧ��sheetName�������κδ���
    If sheetIndex > 0 Then
        
       returnStr = getData4IndexSheet(sheetIndex, exlQsdX, exlQsdY, exlZdX, exlZdY, F1Book1)
       
    Else
    
        returnStr = ""
    
    End If
        
    getData4NameSheet = returnStr
    
End Function


'չ������,�˷�������Ϊָ�����ֵ�sheetҳչ�����ݣ��ʺ��ڶ�sheetҳ��excel
'sheetNameΪҪչ�����ݵ�sheet������
Public Function showData4NameSheet(sheetName As String, showDataStr As String, F1Book1 As F1Book)

    Dim sheetIndex As Integer
    
    '��������sheetNameװ���ɶ�Ӧ��sheetIndex
    sheetIndex = transName2Index(sheetName, F1Book1)
    
    '�������Ϊ0˵��û�ж�Ӧ��sheetName�������κδ���
    If sheetIndex > 0 Then
        
        showData4IndexSheet sheetIndex, showDataStr, F1Book1
    
    End If
        

End Function
'˽�з�������������sheet name װ������ sheet index
'��������ڶ�Ӧ��sheetname������0
Private Function transName2Index(sheetName As String, F1Book1 As F1Book) As Integer

    Dim haveName As Boolean
    haveName = False
    
    Dim sheetNum As Integer
    sheetNum = F1Book1.NumSheets

    For i = 1 To sheetNum
    
        'Action = MsgBox(F1Book1.sheetName(i), vbOKCancel, "ok")

        If F1Book1.sheetName(i) = sheetName Then
        
            'Action = MsgBox(i, vbOKCancel, "ok")
            haveName = True
            
            Exit For
        
        End If
 

    Next i


    '�ж��Ƿ��ж�Ӧ��sheetName��û�еĻ�����0
    If haveName = True Then
        
        transName2Index = i
    Else
        
        transName2Index = 0
        
    End If
    
'Action = MsgBox(transName2Index, vbOKCancel, "ok")

End Function

'˽�з���
'�жϵ�Ԫ���Ƿ�Ϊ��ʽ�true�����ǹ�ʽ�false���ǹ�ʽ��
Private Function haveNoFormula(sheetIndex As Integer, hang As Integer, lie As Integer, F1Book1 As F1Book) As Boolean

Dim s As String

s = F1Book1.FormulaLocalSRC(sheetIndex, hang, lie)

Dim noFormula As Boolean
'���sΪtrue��˵��û�й�ʽ��Ϊfalse˵���й�ʽ
'Action = MsgBox(s, vbOKCancel, "ok")

haveNoFormula = Trim(s) = ""


End Function
Public Function isAllowEdit(isEdit As Boolean, F1Book1 As F1Book)
    F1Book1.AllowInCellEditing = isEdit
    F1Book1.ShowEditBar = isEdit
End Function

Public Function validate4IndexSheet(sheetIndex As Integer, exlQsdX As Integer, exlQsdY As Integer, exlZdX As Integer, exlZdY As Integer) As Boolean
isValidateFailed = False


validate4IndexSheet = Not isValidateFailed
End Function
