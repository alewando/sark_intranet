VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cWX"
Attribute VB_GlobalNameSpace = True
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_sHTML As String


Public Property Get htmlWX() As String
    htmlWX = m_sHTML
End Property


Public Sub OpenURL(sURL As String)

    Dim hFile As Long
    Dim hInet As Long
    Dim lBytesRead As Long
    Dim lResult As Long
    Dim bBuffer() As Byte
    Dim sBuffer As String

    hInet = InternetOpen("OpenURL", _
        INTERNET_OPEN_TYPE_PRECONFIG, vbNullString, _
        vbNullString, 0)
    If hInet Then
        hFile = InternetOpenUrl(hInet, sURL, _
            vbNullString, 0&, _
            INTERNET_FLAG_DONT_CACHE, 0)
        If hFile Then
            ReDim bBuffer(1 To 4096) As Byte
            Do
                lResult = InternetReadFile(hFile, _
                    bBuffer(1), UBound(bBuffer), lBytesRead)
                If lResult Then
                    If lBytesRead Then
                        ' Copy actual number of bytes read
                        ' from byte array onto end of string
                        ' buffer
                        sBuffer = sBuffer & _
                            Left$(StrConv(bBuffer, _
                            vbUnicode), lBytesRead)
                    Else
                        Exit Do
                    End If
                End If
            Loop While lResult
            lResult = InternetCloseHandle(hFile)
        End If
        
        lResult = InternetCloseHandle(hInet)
    End If

    m_sHTML = sBuffer
End Sub
