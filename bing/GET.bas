Attribute VB_Name = "Module1"
Option Explicit
Private Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
'��Ӧÿ��һͼ��ȡ����
Public Function getHtmlStr$(strUrl$)  '��ȡԴ��
Dim a As String
Dim b As Long
Dim c As Long
Dim d As Variant
On Error Resume Next
Dim XmlHttp As Object, stime As String, ntime As String
Set XmlHttp = CreateObject("Microsoft.XMLHTTP")
XmlHttp.open "GET", strUrl, True
XmlHttp.send
stime = Now '��ȡ��ǰʱ��
While XmlHttp.ReadyState <> 4
DoEvents
ntime = Now '��ȡѭ��ʱ��
If DateDiff("s", stime, ntime) > 3 Then getHtmlStr = "": Exit Function '�жϳ���3�뼴��ʱ�˳�����
Wend
a = StrConv(XmlHttp.responseBody, vbUnicode)
b = InStr(1, a, "url") + 6
c = InStr(1, a, "urlbase") - 3
d = URLDownloadToFile(0, "https://cn.bing.com" + Mid(a, b, c - b), App.Path + "/day.jpg", 0, 0)
getHtmlStr = StrConv(d, vbUnicode)
Set XmlHttp = Nothing
End Function
