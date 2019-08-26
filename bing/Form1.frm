VERSION 5.00
Begin VB.Form Form1 
   ClientHeight    =   120
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2025
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   120
   ScaleWidth      =   2025
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Declare Function FindWindowEx Lib "user32" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
'=======================================================================================================================
Dim a As Long
Dim road As String
'========================================================================================================================


Private Sub Form_Load()
Form1.Hide
road = App.Path & "\day.jpg"
If Dir(road) <> "" Then
Kill road
End If
a = getHtmlStr("http://cn.bing.com/HPImageArchive.aspx?format=js&idx=0&n=1") '获取图片
 If a = 0 Then
 Call change '换壁纸为每日一图
 Else
 MsgBox "ERROR图片更换错误"
 End If

 
'=========================================================================================================================
'防止程序多开
If App.PrevInstance Then '如果重复启动
Dim Hwd As Long, mhwnd As Long, HwdOld As Long
Dim i As Integer, Title As String, ClassName As String
mhwnd = Me.hwnd
ClassName = GetWinClass(mhwnd)
Title = GetWinText(mhwnd)
HwdOld = mhwnd
mhwnd = 0
Hwd = FindWindowEx(mhwnd, 0, ClassName, Title)
Do Until Hwd = 0
If Hwd <> HwdOld Then '找到先前启动的本程序
'ShowWindow Hwd, 1
'SetForegroundWindow Hwd
Exit Do
End If
Hwd = FindWindowEx(mhwnd, Hwd, ClassName, Title)
Loop
End
End If
'=====================================================================================================
End Sub
