<!--
Developer:	Barry Behrmann
History:	12/13/1999 - Modified Kevin Dill's email code to send pages.
			11/12/1998 - Email feedback to a specific person.
			12/09/1998 - Moved most of email code into script.asp function.
-->


<!-- #include file="script.asp" -->

<%
if request.form("send")="yes" then
	dim BodyTxt
	BodyTxt = request.form("body")
	if request.form("footer") <> "" then BodyTxt = BodyTxt & NL & NL & NL & "***  sent from " & request("footer")
	if request.form("src")<>"" then BodyTxt = BodyTxt & NL & "***  path = " & mid(request.form("src"), 29)
%>

<html>
<head>
<title>Send Page</title>

<script language=javascript>
<!--
function closeWin(){i = setTimeout("window.close()", 700)}
// -->
</script>

</head>

<body bgcolor=silver onLoad="closeWin()">
<br><center><font size=1 face="ms sans serif, arial, geneva">Sending page...</font></center>

<%
on error resume next
SendMail request.form("subject"), request.form("recipient"), "", BodyTxt
%>

<center><font size=4 face="ms sans serif, arial, geneva"><br><br><b>

<%if err = 0 then%>
	Page sent.
<%else
	err.clear
%>
	<font color=red>Error sending page!</font>
<%end if%>

</b></font></center></body></html>


<%else%>


<html>
<head>
<title>Send Page</title>
<!-- #include file="style.htm" -->

<script language=javascript>
<!--
function SendMail(){
	var valid = true
<%if request("editto") = "yes" then%>
	if (document.frmInfo.recipient.value == ""){
		alert("Please enter recipient...")
		document.frmInfo.recipient.focus()
		valid = false
		}
<%end if%>
	if (valid && (document.frmInfo.subject.value == "")){
		alert("Please enter the subject of the message...")
		document.frmInfo.subject.focus()
		valid = false
		}
	if (valid && (document.frmInfo.body.value == "")){
		alert("Please enter your message...")
		document.frmInfo.body.focus()
		valid = false
		}
	if (valid) document.frmInfo.submit()
	}
// -->
</script>

</head>

<body bgcolor=silver onLoad="document.frmInfo.<%if request("recipient") = "" then%>recipient<%else if request("subject")="" then%>subject<%else%>body<%end if%><%end if%>.focus()"><center>

<form name=frmInfo method=post action="email.asp">
<table border=0>
<%if request("editto") = "yes" then%>
	<tr>
		<td>&nbsp;</td>
	</tr>
<%end if%>
	<tr>
		<td align=right><font size=1 color=navy face="ms sans serif, arial, geneva">To:</font></td>
		<td>
<%if request("editto") = "yes" then%>
			<input type=text name=recipient size=54 maxlength=150 value="<%=request("recipient")%>">
<%else%>
			<font size=1 face="ms sans serif, arial, geneva"><b><%=request("recipient")%></b></font>
			<input type=hidden name=recipient value="<%=request("recipient")%>">
<%end if%>
		</td>
	</tr>
	<tr>
		<td align=right><font size=1 color=navy face="ms sans serif, arial, geneva">Subject:</font></td>
		<td><input type=text name=subject size=54 maxlength=70 value="<%=request("subject")%>"></td>
	</tr>
	<tr>
		<td>&nbsp;</td><td><textarea name=body rows="8" cols="50"></textarea></td>
	</tr>
</table>

<input type=button class=button value="Send" onClick="SendMail()">&nbsp;&nbsp;<input type=button class=button value="Cancel" onClick="window.close()">

<script language=javascript>
<!--
if (window.opener) document.write("<input type=hidden name=src value='" + window.opener.location.href + "'>")
// -->
</script>

<input type=hidden name=footer value="<%=request("footer")%>">
<input type=hidden name=send value="yes">
</form>

</center></body></html>


<%end if%>
