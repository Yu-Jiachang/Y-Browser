VERSION 5.00
Begin VB.MDIForm MDIfrm 
   BackColor       =   &H8000000C&
   Caption         =   "Y浏览器"
   ClientHeight    =   6645
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   13200
   Icon            =   "MDIfrm.frx":0000
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  '窗口缺省
   WindowState     =   2  'Maximized
   Begin VB.Menu Tabs 
      Caption         =   "标签页(&T)"
      WindowList      =   -1  'True
      Begin VB.Menu NewTab 
         Caption         =   "新建标签页"
         Shortcut        =   ^T
      End
      Begin VB.Menu Exit 
         Caption         =   "退出"
         Shortcut        =   ^E
      End
   End
   Begin VB.Menu Settings 
      Caption         =   "设置(&S)"
      Begin VB.Menu Internet_Options 
         Caption         =   "Internet 选项"
         Shortcut        =   ^O
      End
   End
   Begin VB.Menu Help 
      Caption         =   "帮助(&H)"
      Begin VB.Menu GitHub_URL 
         Caption         =   "GitHub 开源地址"
         Shortcut        =   {F2}
      End
      Begin VB.Menu About 
         Caption         =   "关于"
         Shortcut        =   {F1}
      End
   End
End
Attribute VB_Name = "MDIfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub ShellAbout Lib "shell32.dll" Alias "ShellAboutA" (ByVal hwnd As Long, ByVal Caption As String, ByVal OtherStuff As String, ByVal Icon As Long)

Private Sub MDIForm_Load()
    If Command = "" Then
        frmBrowser.brwWebBrowser.GoHome
    Else
        frmBrowser.brwWebBrowser.Navigate Command
    End If
End Sub

Private Sub MDIForm_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If MsgBox("这样会放弃本窗口打开的所有页面，您真的要退出吗？", vbOKCancel + vbQuestion) = vbOK Then
        CreateObject("WScript.Shell").RegDelete "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION\" & App.EXEName + ".exe"
    Else
        Cancel = True
    End If
End Sub

Private Sub NewTab_Click()
    Dim frmBrowser As New frmBrowser
    frmBrowser.Show
    frmBrowser.brwWebBrowser.GoHome
End Sub

Private Sub Exit_Click()
    Unload Me
End Sub

Private Sub Internet_Options_Click()
    Shell "rundll32.exe shell32.dll,Control_RunDLL inetcpl.cpl"
End Sub

Private Sub GitHub_URL_Click()
    Dim frmBrowser As New frmBrowser
    frmBrowser.Show
    frmBrowser.brwWebBrowser.Navigate "https://github.com/VB6-MrYu/Y-Browser"
End Sub

Private Sub About_Click()
    ShellAbout Me.hwnd, App.ProductName & "V" & App.Major & App.Minor, "一款用VB6编写的简单浏览器。" & vbNewLine & _
    "当前内核版本：Microsoft Internet Explorer " & IEVersion, Me.Icon
End Sub
