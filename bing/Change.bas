Attribute VB_Name = "Module3"
Option Explicit

'此修改桌面壁纸为指定图片函数
'==========================================================
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, lpvParam As Any, ByVal fuWinIni As Long) As Long
Private Const SPI_SETDESKWALLPAPER = 20
Private Const SPIF_UPDATEINIFILE = &H1

Public Function change()
    Dim t As Long
    Dim img As String
    img = App.Path & "/day.jpg"
    t = SystemParametersInfo(ByVal SPI_SETDESKWALLPAPER, True, ByVal img, SPIF_UPDATEINIFILE)
End Function
