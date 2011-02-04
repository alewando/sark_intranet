Attribute VB_Name = "startup"
Option Explicit

Private Sub Main()
    Dim objCnet As cnet
    Dim strHTML As String

    Set objCnet = New cnet
    
    strHTML = objCnet.getCnet
    
    'Parses out Cnet news
    If InStr(strHTML, "Today") Then
        strHTML = Right(strHTML, Len(strHTML) - InStr(strHTML, "<b>Today"))
        strHTML = Left(strHTML, InStr(strHTML, "<p>") - 1)
        strHTML = Replace(strHTML, "<font size='-1'>", "<font size='1'>")
        strHTML = Right(strHTML, Len(strHTML) - InStr(strHTML, "size='1'>") + 7)
        strHTML = Replace(strHTML, "<ul>", "")
        strHTML = Replace(strHTML, "<li>", "&nbsp&nbsp&nbsp&nbsp")
        strHTML = Replace(strHTML, "</a>", "</a><br>")
        strHTML = Replace(strHTML, "</ul>", "")
        strHTML = Replace(strHTML, "/news", "http://news.cnet.com/news")
        strHTML = Replace(strHTML, """>", """ target=""_blank"">")
        strHTML = Replace(strHTML, """ >", """ target=""_blank"">") & vbCrLf
    Else
        strHTML = ""
    End If
    
    Open App.Path & "\cnet.txt" For Output As #1
       Write #1, strHTML
    Close #1
    
End Sub

