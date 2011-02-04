<!-- #include file="../../script.asp" -->


<%
'---------------------------------------------------------------------------------
' Description:	Adds a new project.
' History:		06/10/1999 - KDILL - Created
'---------------------------------------------------------------------------------
%>


<html>
<head>
<title>Edit client info</title>
<!-- #include file="../../style.htm" -->
</head>


<body bgcolor=silver><font face="ms sans serif, arial, geneva">

<center>

<%
if request("action") = "edit" then
	
	' UPDATE EXISTING CLIENT...
	WebSite = Clean(request("WebSite"))
	if InStr(WebSite, "://") = 0 then WebSite = "http://" & WebSite
	sql = "UPDATE Client SET "
	sql = sql & "ClientName = '" & Clean(request("ClientName")) & "'"
	sql = sql & ", WebSite = '" & WebSite & "'"
	sql = sql & ", Description = '" & Clean(request("Description")) & "'"
	sql = sql & " WHERE Client_ID = " & request("ClientID")
	DBQuery(sql)
%>

<font size=3>Client info saved...</font>

<script language=javascript>
<!--
var browserAbbr = "";
var browserName = navigator.appName
var browserVersion = parseInt(navigator.appVersion.substring(0,1))
if (browserName=="Microsoft Internet Explorer") {browserAbbr = "IE";}
else if (browserName=="Netscape") {browserAbbr = "NS";}
function closeWin(){
	if (self.opener && ((browserAbbr=="IE" && browserVersion > 3) || (browserAbbr=="NS" && browserVersion > 2))) self.opener.location.reload()
	window.close()
	}
i = setTimeout("closeWin()", 700)
// -->
</script>

<%
else
%>

<script language=javascript>
<!--
function validateFields(){
	return true
	}
// -->
</script>


<form name="frmInfo" method="post" action="details_client.asp" onSubmit="return (validateFields())">
<input type=hidden name=clientid value="<%=request("ClientID")%>">

<%
set rs = DBQuery("select ClientName, WebSite, Description from Client where Client_ID = " & request("ClientID"))
Website = "http://"
if not rs.eof then
	ClientName = rs("ClientName")
	Website = rs("WebSite")
	set Description = rs("Description")
	end if
%>

<table border=0>

	<tr>
		<td valign=middle><font size=1 face="ms sans serif, arial, geneva">
			Client Name:
			</font></td>
		<td><input type=text name="ClientName" value="<% =ClientName %>" size=40 maxlength=100></td>
	</tr>

	<tr>
		<td valign=top><font size=1 face="ms sans serif, arial, geneva">
			Description:
			</font></td>
		<td><textarea name="Description" cols=55 rows=11><% =Description %></textarea></td>
	</tr>

	<tr>
		<td valign=middle><font size=1 face="ms sans serif, arial, geneva">
			Web site:
			</font></td>
		<td><input type=text name="WebSite" value="<% =WebSite %>" size=55 maxlength=100></td>
	</tr>

</table>

<br>
<input type=hidden name=action value="edit">
<input type=submit class=button value="Save client info">
<input type=button class=button value="Cancel" onClick="window.close()">
</form>

<%
rs.close
DataConn.close
set DataConn = nothing
end if
%>

</font></center>


</body>
</html>
