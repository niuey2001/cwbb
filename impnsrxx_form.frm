VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form impnsrxx_form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "数据恢复"
   ClientHeight    =   3090
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7530
   Icon            =   "impnsrxx_form.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3090
   ScaleWidth      =   7530
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0C0C0&
      Height          =   3105
      Left            =   0
      ScaleHeight     =   3045
      ScaleWidth      =   7425
      TabIndex        =   0
      Top             =   0
      Width           =   7485
      Begin VB.TextBox text_file_path 
         BackColor       =   &H00C0FFFF&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
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
         TabIndex        =   4
         Top             =   480
         Width           =   3675
      End
      Begin VB.CommandButton choose_file 
         Caption         =   "选择..."
         Height          =   495
         Left            =   5520
         TabIndex        =   3
         Top             =   480
         Width           =   1455
      End
      Begin VB.CommandButton b_import_nsrxx 
         Caption         =   "恢复"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   3600
         TabIndex        =   2
         Top             =   1680
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "退出"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   5520
         TabIndex        =   1
         Top             =   1680
         Width           =   1335
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "文件路径："
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   240
         TabIndex        =   5
         Top             =   600
         Width           =   1275
      End
   End
   Begin MSComDlg.CommonDialog dia 
      Left            =   2160
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "impnsrxx_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim B() As Byte
Dim PassWord As String
Dim B1 As Byte
Dim i As Long, l As Long, j As Long
Private Sub b_import_nsrxx_Click()
   Dim choose As Integer
     If Me.text_file_path.Text = "" Then
         MsgBox ("请选择恢复的文件")
            Else
                    choose = MsgBox("你确定要恢复到以前数据吗?", vbOKCancel)
                    If choose = 1 Then
                
                  
                    
                    cn.Close
                    FileCopy Me.text_file_path.Text, App.Path & "\" & "financialForm.mdb"
                    cn.Open
                    MsgBox ("已经成功恢复")
                   
                    Else
                       ' Exit Sub
                    End If
End If



End Sub

Private Sub choose_file_Click()
Dim fileName As String
On Error GoTo errpro
Me.dia.InitDir = App.Path

Me.dia.Filter = "mdb文件(*.mdb)|*.mdb"


Me.dia.ShowOpen

fileName = Me.dia.fileName
If fileName = "" Then
   GoTo errpro
Else
   Me.text_file_path.Text = fileName
   Exit Sub
End If
errpro:
'MsgBox "你没有选择任何文件、文件不存在或文件已作废。", vbCritical, "选择错误"
End Sub


Private Sub Command3_Click()

End Sub

Private Sub Command1_Click()
Unload Me
End Sub
