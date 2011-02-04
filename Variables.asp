<!--
Developer:    Kevin Dill
Date:         08/17/1998
Description:  Main welcome page first displayed to user.
-->


<!-- #include file="header.asp" -->

<%
On Error Resume Next

Response.Write("<h1>Application Variables</h1><br><hl>" & vbCRLF)
For each variable in Application.Contents
	Response.Write("<B>" & variable & ": </B>")
	Response.Write(Application(variable) & "<BR>" & vbCRLF)
Next

Response.Write("<BR><h1>Session Variables</h1><br><hl>" & vbCRLF)
For each variable in Session.Contents
	Response.Write("<B>" & variable & ": </B>")
	Response.Write(Session(variable) & "<BR>" & vbCRLF)
Next

'Response.Write("<BR><h1>Request Variables</h1><br><hl>" & vbCRLF)
'For each variable in Request.Servervariables
'	Response.Write("<B>" & variable & ": </B>")
'	Response.Write(Request.Servervariables(variable) & "<BR>" & vbCRLF)
'Next

%>


<!-- #include file="footer.asp" -->