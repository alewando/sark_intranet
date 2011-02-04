<HTML><head><title>Cincinnati Weather</title></head>
<body>

<b>WEATHER</b><br><br>
<table>
<tr><%dim objFile, objFSOdim strHTMLset objFSO = server.CreateObject("Scripting.FileSystemObject")objFSO.GetFile ("c:\inetpub\wwwroot\content\weather.txt")Set objFile = objFSO.OpenTextFile("c:\inetpub\wwwroot\content\weather.txt")strHTML = objFile.readallstrHTML = replace(strHTML,chr(34),"")objFile.closeset objFile = nothingset objFSO = nothing

Response.Write (strHTML)%><tr>
</table>

<br><br><br><b>CNET</b><br><br><%set objFSO = server.CreateObject("Scripting.FileSystemObject")objFSO.GetFile ("c:\inetpub\wwwroot\content\cnet.txt")Set objFile = objFSO.OpenTextFile("c:\inetpub\wwwroot\content\cnet.txt")
strHTML = objFile.readallstrHTML = replace(strHTML,chr(34),"")objFile.closeset objFile = nothingset objFSO = nothing
Response.Write (strHTML)%></font>
<br><br>
<b>401K</b><br><br><table border=1 cellpadding=0 cellspacing=0><tr><td>
<table border=0 cellpadding=2 cellspacing=0 bgcolor=#d0d0d0 width=115>
<%set objFSO = server.CreateObject("Scripting.FileSystemObject")objFSO.GetFile ("c:\inetpub\wwwroot\content\stock.txt")Set objFile = objFSO.OpenTextFile("c:\inetpub\wwwroot\content\stock.txt")
strHTML = objFile.readallstrHTML = replace(strHTML,chr(34),"")objFile.closeset objFile = nothingset objFSO = nothing
Response.Write (strHTML)%>
</table>
</td></tr>
</table></body></html>