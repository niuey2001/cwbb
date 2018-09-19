Attribute VB_Name = "openPath"
Private Declare Function SHBrowseForFolder Lib "shell32.dll" Alias "SHBrowseForFolderA" (LpBrowseInfo As BROWSEINFO) As Long
Private Declare Function SHGetPathFromIDlist Lib "shell32.dll" Alias "SHGetPathFromIDListA" (ByVal pidl As Long, ByVal pszPath As String) As Long
Private Type BROWSEINFO
  hOwner As Long
  pidlroot As Long
  pszDisplayName As String
  lpszTitle As String
  ulFlags As Long
  lpfn As Long
  lparam As Long
  iImage As Long
End Type
Public Function GetFolder(Optional Title As String, Optional hwnd) As String
    Dim bi As BROWSEINFO
    Dim pidl As Long
    Dim folder As String
     
folder = Space(255)
     
With bi
   If IsNumeric(hwnd) Then .hOwner = hwnd
   .ulFlags = BIF_RETURNONLYFSDIRS
   .pidlroot = 0
   If Title <> "" Then
      .lpszTitle = Title & Chr$(0)
   Else
      .lpszTitle = "选择目录" & Chr$(0)
    End If
End With

pidl = SHBrowseForFolder(bi)
If SHGetPathFromIDlist(ByVal pidl, ByVal folder) Then
    GetFolder = Left(folder, InStr(folder, Chr$(0)) - 1)
Else
    GetFolder = ""
End If

End Function
'调用方式
'FilePath=GetFolder("打开一个目录", Form1.hwnd)
