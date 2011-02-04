<%
set fso = CreateObject("Scripting.FileSystemObject")
WebRootDir = Application("WebRootDir")

if isnumeric(request("slide")) then
	slidenum = request("slide")
	filepath = WebRootDir & Application("web") & "\extras\" & request("dir") & "\Slide" & slidenum
	if not fso.FileExists(filepath & ".gif") and not fso.FileExists(filepath & ".jpg") then slidenum=1
else
	slidenum = 1
end if
%>

<HTML>
<HEAD>
<TITLE><%=request("title")%> Presentation (Slide <%=slidenum%>)</TITLE>
</HEAD>


<BODY bgcolor=black>
<font size=1 face=arial color=white><center>

Click on the slide to view the next one...<br>

<%
extension = ".gif"
filepath = WebRootDir & Application("web") & "\extras\" & request("dir") & "\Slide" & slidenum & ".gif"

if not fso.FileExists(filepath) then extension = ".jpg"
%>

<a href="slideshow.asp?title=<%=Server.URLEncode(request("title"))%>&dir=<%=Server.URLEncode(request("dir"))%>&slide=<%=slidenum+1%>">
<img src="/<%=Application("web") & "/extras/" & request("dir") & "/slide" & slidenum & extension%>" border=0>
</a>


</center></font>
</BODY>
</HTML>
