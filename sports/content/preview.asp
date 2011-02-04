 <%@ Language = Vbscript%>
<%
 title = Request("Title")
 text = Request("Text")
%>

<HTML>
<HEAD>
 <TITLE>Preview: <%=title%></TITLE>
<!-- #include file="../../style.htm" --> 
</HEAD>
<BODY BGCOLOR="gray">
<TABLE CLASS=tableShadowSilver WIDTH=100%>
<TR><TD>
<B><%=title%>:</B><BR>
<DIV CLASS=sportDescription>
<%=text%>
</DIV>
</TD></TR>
</TABLE>
<CENTER>
 <INPUT TYPE=BUTTON CLASS=button VALUE="Close" onClick="window.close()">
</CENTER>
</BODY>
</HTML>