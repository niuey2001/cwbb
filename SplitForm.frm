VERSION 5.00
Begin VB.Form frmSplittable 
   Caption         =   "Splittable Child Window"
   ClientHeight    =   5490
   ClientLeft      =   4065
   ClientTop       =   2175
   ClientWidth     =   9060
   Icon            =   "fSplit.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   5490
   ScaleWidth      =   9060
   Begin VB.ListBox lstFiles 
      Height          =   1260
      IntegralHeight  =   0   'False
      Left            =   60
      TabIndex        =   4
      Top             =   3900
      Width           =   8895
   End
   Begin VB.TextBox Text2 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3675
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Text            =   "fSplit.frx":014A
      Top             =   120
      Width           =   6555
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3675
      Left            =   60
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Text            =   "fSplit.frx":0150
      Top             =   120
      Width           =   2235
   End
   Begin VB.PictureBox picSplit 
      Height          =   4155
      Left            =   2220
      MousePointer    =   9  'Size W E
      ScaleHeight     =   4095
      ScaleWidth      =   195
      TabIndex        =   2
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox picVSplit 
      Height          =   255
      Left            =   0
      ScaleHeight     =   195
      ScaleWidth      =   9015
      TabIndex        =   3
      Top             =   3720
      Width           =   9075
   End
End
Attribute VB_Name = "frmSplittable"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim cHS As New cSplitDDC
Dim cVS As New cSplitDDC

Private Function GetFileText(ByVal sFile As String, ByRef sText As String) As Boolean
Dim iFIle As Integer, lLen As Long
    iFIle = FreeFile
    On Error Resume Next
    Open sFile For Binary Access Read As #iFIle
    If (Err.Number = 0) Then
        lLen = LOF(iFIle)
        sText = String$(lLen, 0)
        Get #iFIle, , sText
        If (Err.Number = 0) Then
            GetFileText = True
        End If
        Close #iFIle
    End If

End Function

Private Sub Form_Load()

    With cHS
        .Orientation = espVertical
        .Border(espbLeft) = 32
        .Border(espbRight) = 32
        .SplitObject = picSplit
    End With
    With cVS
        .Orientation = espHorizontal
        .Border(espbBottom) = 32
        .Border(espbTop) = 64
        .Border(espbLeft) = 2
        .Border(espbRight) = 2
        .SplitObject = picVSplit
    End With
    
    Dim sText As String, sFile As String
    sFile = App.Path & "\mfrmMain.frm"
    If (GetFileText(sFile, sText)) Then
        Text1.Text = sText
    Else
        Text1.Text = "Source code to file '" & sFile & "' could not be found."
    End If
    sFile = App.Path & "\cSplitDC.cls"
    If (GetFileText(sFile, sText)) Then
        Text2.Text = sText
    Else
        Text2.Text = "Source code to file '" & sFile & "' could not be found."
    End If
    
    sFile = Dir(App.Path & "\*.*")
    Do While Len(sFile) > 0
        lstFiles.AddItem sFile
        sFile = Dir
    Loop
    
    If (Me.ScaleHeight \ 4 * 3 >= cVS.Border(espbTop) * Screen.TwipsPerPixelY) Then
        picVSplit.Top = Me.ScaleHeight \ 4 * 3
    Else
        picVSplit.Top = cVS.Border(espbTop) * Screen.TwipsPerPixelY
    End If
    
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cHS.SplitterFormMouseMove X, Y
    cVS.SplitterFormMouseMove X, Y
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If (cHS.SplitterFormMouseUp(X, Y)) Then
        Form_Resize
    ElseIf (cVS.SplitterFormMouseUp(X, Y)) Then
        Form_Resize
    End If
    
End Sub

Private Sub Form_Resize()
Dim lH As Long

    lH = picVSplit.Top + 2 * Screen.TwipsPerPixelY
    With Text1
        .Move .Left, .Top, _
            picSplit.Left - .Left - 4 * Screen.TwipsPerPixelX, _
            lH
    End With
    Text2.Move picSplit.Left + 2 * Screen.TwipsPerPixelX, Text1.Top, _
        (Me.ScaleWidth - picSplit.Left - 4 * Screen.TwipsPerPixelX - Text1.Left), _
        lH
    With lstFiles
        .Move .Left, picVSplit.Top + picVSplit.Height - 2 * Screen.TwipsPerPixelY, _
            Me.ScaleWidth - .Left * 2, Me.ScaleHeight - (picVSplit.Top + picVSplit.Height - 2 * Screen.TwipsPerPixelY) - 2 * Screen.TwipsPerPixelY
    End With
    cHS.Border(espbBottom) = (Me.ScaleHeight - lstFiles.Top) \ Screen.TwipsPerPixelY
    
End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cHS.SplitterMouseDown Me.hWnd, X, Y
End Sub

Private Sub picVSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    cVS.SplitterMouseDown Me.hWnd, X, Y
End Sub
