Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
'必应每日一图获取函数
Public Function getHtmlStr$(strUrl$)  '获取源码
Dim a As String
Dim b As Long
Dim c As Long
Dim d As Variant
On Error Resume Next
Dim XmlHttp As Object, stime As String, ntime As String
Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
XmlHttp.open "GET", strUrl, True
XmlHttp.send
stime = Now '获取当前时间
While XmlHttp.ReadyState <> 4
DoEvents
ntime = Now '获取循环时间
If DateDiff("s", stime, ntime) > 3 Then getHtmlStr = "": Exit Function '判断超出3秒即超时退出过程
Wend
a = StrConv(XmlHttp.responseBody, vbUnicode)
b = InStr(1, a, "url") + 6
c = InStr(1, a, "urlbase") - 3
d = URLDownloadToFile(0, "https://cn.bing.com" + Mid(a, b, c - b), App.Path + "/day.jpg", 0, 0)
getHtmlStr = StrConv(d, vbUnicode)
Set XmlHttp = Nothing
End Function
