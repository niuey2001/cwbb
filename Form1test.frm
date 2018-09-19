VERSION 5.00
Object = "{B0475000-7740-11D1-BDC3-0020AF9F8E6E}#6.1#0"; "TTF16.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6705
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8160
   LinkTopic       =   "Form1"
   ScaleHeight     =   6705
   ScaleWidth      =   8160
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   1095
      Left            =   5400
      TabIndex        =   1
      Top             =   5040
      Width           =   2415
   End
   Begin TTF160Ctl.F1Book F1Book1 
      Height          =   3615
      Left            =   2640
      TabIndex        =   0
      Top             =   960
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6376
      _0              =   $"Form1test.frx":0000
      _1              =   $"Form1test.frx":0409
      _2              =   $"Form1test.frx":0812
      _3              =   $"Form1test.frx":0C1B
      _4              =   $"Form1test.frx":1024
      _count          =   5
      _ver            =   2
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
getUrl App.Path & "\" & "资产负债表.xls", Me.F1Book1
End Sub
