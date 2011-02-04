Attribute VB_Name = "startup"
Option Explicit

Private Sub Main()
    Dim objStock As stock
    Dim strHTML As String
    Dim i As Integer

    Set objStock = New stock
    
    strHTML = objStock.getStock

    strHTML = Right(strHTML, Len(strHTML) - (InStr(strHTML, "<TR ALIGN=CENTER VALIGN=CENTER>")) + 1)
    strHTML = Left(strHTML, (InStr(strHTML, "</TABLE>") - 1))
    strHTML = Replace(strHTML, "<TR ALIGN=CENTER VALIGN=CENTER>", "<tr>")
    strHTML = Replace(strHTML, "WIDTH=70 ALIGN=LEFT class=cBW", "align=center><font size=1 face=arial")
    strHTML = Replace(strHTML, Chr(34), "")
    strHTML = Replace(strHTML, "..", "http://quicken.excite.com/investments")
    strHTML = Replace(strHTML, "ALIGN=LEFT class=cBW>&nbsp;&nbsp;<B", "align=right><font size=1 face=arial")
    strHTML = Replace(strHTML, "&nbsp</B>", "</font>")
    strHTML = Replace(strHTML, "<TD class=cBW>N/A</TD>", "")
    strHTML = Replace(strHTML, "<TD class=cBW>MUTUAL</TD>", "")
    strHTML = Replace(strHTML, "class=cBW><FONT COLOR=#FF0000><B", "align=right><FONT size=1 FACE=arial COLOR=#FF0000")
    strHTML = Replace(strHTML, "class=cBW><FONT COLOR=#008800><B", "align=right><FONT size=1 FACE=arial COLOR=#008800")
    strHTML = Replace(strHTML, "class=cBW><B>0", "align=right><FONT size=1 FACE=arial COLOR=gray>0.00")
    
    For i = 1 To 7
        strHTML = Left(strHTML, (InStr(strHTML, "<TD class=cBW><SPAN") - 1)) & Right(strHTML, (Len(strHTML) - (InStr(strHTML, "<TD class=cBW><SPAN") + 49)))
        strHTML = Left(strHTML, (InStr(strHTML, "<NOBR>") - 2)) & Right(strHTML, (Len(strHTML) - (InStr(strHTML, "</B></NOBR>") + 10)))
    Next
    
    strHTML = Left(strHTML, Len(strHTML) - 7)
    strHTML = Replace(strHTML, "<A", "<a target=_blank")
    
    Open App.Path & "\stock.txt" For Output As #1
       Print #1, strHTML
    Close #1
    
End Sub


