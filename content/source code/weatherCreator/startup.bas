Attribute VB_Name = "startup"
Option Explicit

Private Sub Main()
    Dim objWeather As weather
    Dim strHTML As String

    Set objWeather = New weather
    
    'This parses out the 7-day forecast table
    strHTML = objWeather.getWeather
    strHTML = Replace(strHTML, Chr(34), "")
    strHTML = Right(strHTML, Len(strHTML) - (InStr(strHTML, "<!-- forecast include -->")))
    strHTML = Right(strHTML, Len(strHTML) - (InStr(strHTML, "<TABLE") - 1))
    strHTML = Left(strHTML, InStr(strHTML, "</TABLE>") + 7)
    
    'This transforms it to our look and feel
    strHTML = Right(strHTML, Len(strHTML) - (InStr(strHTML, "<TD ALIGN=center VALIGN=middle WIDTH=66 BGCOLOR=#649CCC><FONT FACE=Arial, Helvetica, Chicago, Sans Serif SIZE=2 COLOR=#FFFFFF>") - 1))
    strHTML = Replace(strHTML, "  ", "")
    strHTML = Replace(strHTML, "<TD ALIGN=center VALIGN=middle WIDTH=66 BGCOLOR=#649CCC><FONT FACE=Arial, Helvetica, Chicago, Sans Serif SIZE=2 COLOR=#FFFFFF><B>", "<td align=center class=""sectionBody"" nowrap>")
    strHTML = Replace(strHTML, "</B></FONT></TD>", "<br>")
    strHTML = Replace(strHTML, "<TD ALIGN=center VALIGN=middle WIDTH=65 BGCOLOR=#E4ECF4>", "")
    strHTML = Replace(strHTML, "WIDTH=52 HEIGHT=52", "WIDTH=26 HEIGHT=26")
    strHTML = Replace(strHTML, "></TD>" & Chr(10) & "<FONT FACE=Arial, Helvetica, Chicago, Sans Serif SIZE=2>", "" & Chr(39))
    strHTML = Replace(strHTML, "</FONT></TD>", Chr(39) & "><br>")
    strHTML = Replace(strHTML, "<TD ALIGN=right VALIGN=middle WIDTH=65 BGCOLOR=#E4ECF4><FONT FACE=Arial, Helvetica, Chicago, Sans Serif SIZE=2>", "")
    strHTML = Replace(strHTML, "<TR>", "")
    strHTML = Replace(strHTML, "</TR>", "")
    strHTML = Replace(strHTML, "</TABLE>", "")
    strHTML = Replace(strHTML, "</FONT>'><br>", "</td>")
    strHTML = Replace(strHTML, "<TD ALIGN=center VALIGN=top HEIGHT=2><IMG WIDTH=6 HEIGHT=2 SRC=http://image.weather.com/pics/blank.gif ALT=></TD>", "")
    
    Open App.Path & "\weather.txt" For Output As #1
       Write #1, strHTML
    Close #1
    
End Sub



