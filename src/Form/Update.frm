VERSION 5.00
Object = "{EAB22AC0-30C1-11CF-A7EB-0000C05BAE0B}#1.1#0"; "ieframe.dll"
Begin VB.Form Update 
   Caption         =   "检查更新"
   ClientHeight    =   5625
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13500
   LinkTopic       =   "Form1"
   ScaleHeight     =   5625
   ScaleWidth      =   13500
   StartUpPosition =   3  '窗口缺省
   Begin SHDocVwCtl.WebBrowser UpdateWeb 
      Height          =   5655
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   13575
      ExtentX         =   23945
      ExtentY         =   9975
      ViewMode        =   0
      Offline         =   0
      Silent          =   0
      RegisterAsBrowser=   0
      RegisterAsDropTarget=   1
      AutoArrange     =   0   'False
      NoClientEdge    =   0   'False
      AlignLeft       =   0   'False
      NoWebView       =   0   'False
      HideFileNames   =   0   'False
      SingleClick     =   0   'False
      SingleSelection =   0   'False
      NoFolders       =   0   'False
      Transparent     =   0   'False
      ViewID          =   "{0057D0E0-3573-11CF-AE69-08002B2E1262}"
      Location        =   "http:///"
   End
End
Attribute VB_Name = "Update"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Const MF_BYCOMMAND = &H0&
Private Const MF_BYPOSITION = &H400&
Private Const MF_DISABLED = &H2&
Private Const MF_ENABLED = &H0&
Private Const MF_GRAYED = &H1&
Private Declare Function GetSystemMenu Lib "user32" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32" (ByVal hMenu As Long) As Long

Private Function DisableCloseMenu(ByVal dwMenu As Long) As Boolean
Dim hMenu As Long
Dim nCount As Long
DisableCloseMenu = False
hMenu = GetSystemMenu(dwMenu, False)
If hMenu <> 0 Then
nCount = GetMenuItemCount(hMenu)
If RemoveMenu(hMenu, nCount - 1, MF_DISABLED Or MF_BYPOSITION) <> 0 Then
If DrawMenuBar(dwMenu) <> 0 Then
DisableCloseMenu = True
End If
End If
End If
End Function

Private Function EnableCloseMenu(ByVal dwMenu As Long) As Boolean
Dim hMenu As Long
EnableCloseMenu = False
hMenu = GetSystemMenu(dwMenu, True)
If hMenu = 0 Then
If DrawMenuBar(dwMenu) <> 0 Then
EnableCloseMenu = True
End If
End If
End Function

Private Sub Form_Load()
DisableCloseMenu Me.hwnd
UpdateWeb.Navigate "https://apps.yujiachang.linkpc.net/Y-Browser/Update.html?" & App.Major & "." & App.Minor
End Sub

Private Sub Form_Resize()
UpdateWeb.Width = Me.ScaleWidth
UpdateWeb.Height = Me.ScaleHeight
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
End Sub

Private Sub UpdateWeb_WindowClosing(ByVal IsChildWindow As Boolean, Cancel As Boolean)
Cancel = True
EnableCloseMenu Me.hwnd
End Sub
