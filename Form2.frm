VERSION 5.00
Begin VB.Form register_form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�û�ע�ᴰ��"
   ClientHeight    =   6585
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9675
   Icon            =   "Form2.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "Form2.frx":16AC2
   ScaleHeight     =   6585
   ScaleWidth      =   9675
   StartUpPosition =   2  '��Ļ����
   Begin VB.OptionButton user_type 
      BackColor       =   &H00FFC0C0&
      Caption         =   "��ͨ��˰��"
      Height          =   495
      Index           =   0
      Left            =   3480
      TabIndex        =   4
      Top             =   3600
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton b_back 
      BackColor       =   &H00FF8080&
      Caption         =   "����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6480
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox text_password_two 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   3
      Top             =   3075
      Width           =   1695
   End
   Begin VB.CommandButton b_register 
      BackColor       =   &H00FF8080&
      Caption         =   "ע��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   4440
      Width           =   1575
   End
   Begin VB.CommandButton b_clear 
      BackColor       =   &H00FF8080&
      Caption         =   "���"
      BeginProperty Font 
         Name            =   "����"
         Size            =   14.25
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   2160
      MaskColor       =   &H00FF8080&
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox text_username 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   315
      IMEMode         =   3  'DISABLE
      Left            =   3480
      TabIndex        =   0
      Top             =   1620
      Width           =   1695
   End
   Begin VB.TextBox text_password 
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "����"
         Size            =   15
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      IMEMode         =   3  'DISABLE
      Left            =   3480
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2355
      Width           =   1695
   End
   Begin VB.OptionButton user_type 
      BackColor       =   &H00FFC0C0&
      Caption         =   "˰��Ա"
      Height          =   495
      Index           =   1
      Left            =   4800
      TabIndex        =   5
      Top             =   3600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "�û����ͣ�"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   1800
      TabIndex        =   12
      Top             =   3720
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "ȷ�����룺"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   1800
      TabIndex        =   11
      Top             =   3000
      Width           =   1815
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "���û�ע��"
      BeginProperty Font 
         Name            =   "����"
         Size            =   21.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   615
      Left            =   1200
      TabIndex        =   10
      Top             =   360
      Width           =   3975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "�û�����"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   1800
      TabIndex        =   9
      Top             =   1560
      Width           =   1935
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "��   �룺"
      BeginProperty Font 
         Name            =   "����"
         Size            =   15.75
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   615
      Left            =   1800
      TabIndex        =   1
      Top             =   2280
      Width           =   1575
   End
End
Attribute VB_Name = "register_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub b_back_Click()
    login_form.text_name = ""
    login_form.text_password = ""
    register_form.Hide
    login_form.Show
End Sub

Private Sub b_clear_Click()
text_username.Text = ""
text_password.Text = ""
text_password_two.Text = ""
End Sub

Private Sub b_register_Click()
Dim sql As String
Dim userInfoRs As ADODB.Recordset
Dim userNameRs As ADODB.Recordset



Dim v_name As String  '�û���
Dim v_password As String  '����
Dim v_password_two As String  'ȷ������
Dim v_user_type As String
Dim success As Integer



v_name = text_username.Text
v_password = text_password.Text
v_password_two = text_password_two.Text

If user_type(0).value Then
    v_user_type = "1"
Else
v_user_type = "0"
End If

If v_name = "" Then
    MsgBox "�û�������Ϊ�գ�"
    Exit Sub
    
End If
If v_password = "" Then
    MsgBox "���벻��Ϊ�գ�"
    Exit Sub
    
End If

If v_password <> v_password_two Then
    MsgBox "�������벻һ�£�"
    text_password_two.SetFocus
    Exit Sub
End If

'���ݿ����
Call check_condatabase
sql = "select user_name from t_user_info where user_name = '" & v_name & "'"
Set userNameRs = cn.Execute(sql)
If Not userNameRs.EOF Then
    MsgBox "�Բ��𣬴��û����Ѵ��ڣ�"
    Exit Sub
End If
userNameRs.Close
Set userNameRs = Nothing
  
sql = "select * from t_user_info"
Set userInfoRs = New ADODB.Recordset
Set userInfoRs.ActiveConnection = cn
userInfoRs.LockType = adLockOptimistic
userInfoRs.CursorType = adOpenKeyset
userInfoRs.Open sql
  
userInfoRs.AddNew '�������Ʒ��Ϣ
userInfoRs("user_name") = v_name
userInfoRs("password") = v_password
userInfoRs("user_type") = v_user_type
userInfoRs.Update
userInfoRs.Close
Set userInfoRs = Nothing


success = MsgBox("ע��ɹ�����ת����½����", 1, "��ʾ")
If success = 1 Then
    Unload Me
   
    Unload login_form
    login_form.Combo1.Text = v_name
    login_form.Show
    
   
   
End If

















'v_current_path = App.Path
'MsgBox v_current_path
'Set fileObj = CreateObject("Scripting.FileSystemObject")
'Set writeObj = fileObj.CreateTextFile(v_current_path & "\user_info.ini", True)

'writeObj.WriteLine (v_name & "," & v_password)
'writeObj.Close


End Sub


Private Sub Form_Load()
Me.text_password = ""
Me.text_username = ""
Me.text_password_two = ""
user_type(0).value = True
b_register.Default = True
End Sub

