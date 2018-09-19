Attribute VB_Name = "fomularone"
'全局变量
Public isValidateFailed As Boolean       '验证是否通过的标志位，true：为通过；false：通过

'Public isShowData As Boolean
Public xlApp As New excel.Application

Public validateFailedMes As String       '验证失败时存储的错误提示信息，是在vts中定义的信息

Public isGetErrorMes As Boolean          '验证失败事件判断当前触发验证的来源是否为getErrorMes方法

Public validateFailedColor               '验证为通过时，cell显示的颜色，默认为红色
'取得excel模版
Public Function getUrl(theUrl As String, F1Book1 As F1Book)

If Dir(theUrl) = "" Then '文件不存在
    MsgBox "报表不存在！"
Else

 F1Book1.URL = theUrl
 
  
End If
   
End Function
'取得表上的数据和其坐标，并作为string返回,此方法默认从第一个sheet页中取得数据，适合于单sheet的excel
'exlQsdX,exlQsdY为数据的起始点坐标
'exlZdX,exlZdY为数据的终点坐标
Public Function getData(exlQsdX As Integer, exlQsdY As Integer, exlZdX As Integer, exlZdY As Integer, F1Book1 As F1Book) As String
    
    getData = getData4IndexSheet(1, exlQsdX, exlQsdY, exlZdX, exlZdY, F1Book1)

End Function

Public Function showData(showDataStr As String, F1Book1 As F1Book)
showData4IndexSheet 1, showDataStr, F1Book1
End Function

'取得表上的数据和其坐标，并作为string返回,此方从指定的sheet页中取得数据，适合于多sheet页的excel
'sheetIndex为要取得数据的sheet的序号，从左边为1开始计算
'exlQsdX,exlQsdY为数据的起始点坐标
'exlZdX,exlZdY为数据的终点坐标
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
'展现数据,此方法可以为指定序号的sheet页展现数据，适合于多sheet页的excel
'sheetIndex为要展现数据的sheet的序号，从左边为1开始计算


Public Function showData4IndexSheet(sheetIndex As Integer, showDataStr As String, F1Book1 As F1Book)
'isShowData = True

'1 第一次解析字符串，array1()数组里面元素为x,y,value （坐标，值）
Dim array1() As String
array1() = Split(showDataStr, ";")


'2 遍历数组处理每个 坐标，值
Dim array2() As String
For i = 0 To UBound(array1) - 1
    '3 第2次解析字符串,对于每一次循环：array2[0]为横坐标，array2[1]为竖坐标，array2[2]为值
    array2() = Split(array1(i), ",")
    
    
 '判断此单元格有公式  就不赋值   lwy添加
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
           
        '4 为表格符值
    If array2(2) <> "" Then
    F1Book1.EntrySRC(sheetIndex, array2(0), array2(1)) = array2(2)
    'F1Book1.entr
    End If
    

End If



Next i

'isShowData = False
End Function


'取得表上的数据和其坐标，并作为string返回,此方从指定的sheet页中取得数据，适合于多sheet页的excel
'sheetName为要取得数据的sheet页的名字
'exlQsdX,exlQsdY为数据的起始点坐标
'exlZdX,exlZdY为数据的终点坐标
Public Function getData4NameSheet(sheetName As String, exlQsdX As Integer, exlQsdY As Integer, exlZdX As Integer, exlZdY As Integer, F1Book1 As F1Book) As String
    Dim returnStr As String
    Dim sheetIndex As Integer
    
    '将给定的sheetName装换成对应的sheetIndex
    sheetIndex = transName2Index(sheetName, F1Book1)
    
    '如果返回为0说明没有对应的sheetName，则不作任何处理
    If sheetIndex > 0 Then
        
       returnStr = getData4IndexSheet(sheetIndex, exlQsdX, exlQsdY, exlZdX, exlZdY, F1Book1)
       
    Else
    
        returnStr = ""
    
    End If
        
    getData4NameSheet = returnStr
    
End Function


'展现数据,此方法可以为指定名字的sheet页展现数据，适合于多sheet页的excel
'sheetName为要展现数据的sheet的名字
Public Function showData4NameSheet(sheetName As String, showDataStr As String, F1Book1 As F1Book)

    Dim sheetIndex As Integer
    
    '将给定的sheetName装换成对应的sheetIndex
    sheetIndex = transName2Index(sheetName, F1Book1)
    
    '如果返回为0说明没有对应的sheetName，则不作任何处理
    If sheetIndex > 0 Then
        
        showData4IndexSheet sheetIndex, showDataStr, F1Book1
    
    End If
        

End Function
'私有方法，将给定的sheet name 装换成其 sheet index
'如果不存在对应的sheetname，返回0
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


    '判断是否有对应的sheetName，没有的话返回0
    If haveName = True Then
        
        transName2Index = i
    Else
        
        transName2Index = 0
        
    End If
    
'Action = MsgBox(transName2Index, vbOKCancel, "ok")

End Function

'私有方法
'判断单元格是否为公式项。true：不是公式项；false：是公式项
Private Function haveNoFormula(sheetIndex As Integer, hang As Integer, lie As Integer, F1Book1 As F1Book) As Boolean

Dim s As String

s = F1Book1.FormulaLocalSRC(sheetIndex, hang, lie)

Dim noFormula As Boolean
'如果s为true，说明没有公式，为false说明有公式
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
