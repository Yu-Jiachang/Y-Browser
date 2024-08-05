Attribute VB_Name = "modMain"
Dim IEVersion
Dim szIEVersion
Public IEZhuBanBen As Long

Sub Main()
    On Error GoTo NotInstallIE
    IEVersion = CreateObject("wscript.shell").RegRead("HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\Internet Explorer\svcVersion")
    szIEVersion = Split(IEVersion, ".")
    IEZhuBanBen = szIEVersion(0)
    CreateObject("WScript.Shell").RegWrite "HKEY_CURRENT_USER\Software\Microsoft\Internet Explorer\Main\FeatureControl\FEATURE_BROWSER_EMULATION\" & App.EXEName + ".exe", IEZhuBanBen * 1000, "REG_DWORD"
    MDIfrm.Show
    Update.Show vbModal, MDIfrm

Exit Sub
NotInstallIE:
MsgBox "您的计算机上似乎未安装 “Microsoft Internet Explorer” 产品……" & vbNewLine & _
"程序启动失败！", vbCritical
End Sub
