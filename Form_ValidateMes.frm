VERSION 5.00
Begin VB.Form Form_ValidateMes 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Form1"
   ClientHeight    =   5370
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9630
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MDIChild        =   -1  'True
   MinButton       =   0   'False
   ScaleHeight     =   5370
   ScaleWidth      =   9630
   Begin VB.TextBox Text_Wrong_Mes 
      CausesValidation=   0   'False
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   4755
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   600
      Width           =   9615
   End
   Begin VB.Label Label1 
      Caption         =   "¥ÌŒÛ–≈œ¢"
      BeginProperty Font 
         Name            =   "ÀŒÃÂ"
         Size            =   12
         Charset         =   134
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   480
      TabIndex        =   1
      Top             =   120
      Width           =   4815
   End
End
Attribute VB_Name = "Form_ValidateMes"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
