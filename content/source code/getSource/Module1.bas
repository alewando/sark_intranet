Attribute VB_Name = "modInet"
Option Explicit

Public Const INTERNET_OPEN_TYPE_PRECONFIG As Long = 0
Public Const INTERNET_FLAG_DONT_CACHE As _
    Long = &H4000000
Declare Function InternetCloseHandle Lib "wininet" _
    (ByVal hInet As Long) As Long
Declare Function InternetOpen Lib "wininet" Alias _
    "InternetOpenA" (ByVal lpszAgent As String, ByVal _
    dwAccessType As Long, ByVal lpszProxyName As _
    String, ByVal lpszProxyBypass As String, ByVal _
    dwFlags As Long) As Long
Declare Function InternetOpenUrl Lib "wininet" Alias _
    "InternetOpenUrlA" (ByVal hInet As Long, ByVal _
    lpszUrl As String, ByVal lpszHeaders As String, _
    ByVal dwHeadersLength As Long, ByVal dwFlags As _
    Long, ByVal dwContext As Long) As Long
Declare Function InternetReadFile Lib "wininet" _
    (ByVal hFile As Long, lpBuffer As Any, ByVal _
    dwBytesToRead As Long, lpBytesRead As Long) As Long



