<!-- #include file="../section.asp" -->


<%
'---------------------------------------------------------------------------------
' Description:	Email general section.
' History:		11/18/1998 - KDILL - Created
'				12/14/1998 - KDILL - Updated email functionality to login/visit email systems.
'---------------------------------------------------------------------------------
%>


<script language=javascript>
<!--
function EmailPrep(){
	var system = "UNKNOWN"
	if (document.frmInfo.system.options) system = document.frmInfo.system.options[document.frmInfo.system.selectedIndex].value
	document.frmInfo.password.value = document.frmInfo.pass.value
	document.frmInfo.pass.value=''
	setCookie("System", system)
	setCookie("Username", document.frmInfo.username.value)
	}
// -->
</script>

<img src="images/email_letter.gif" width=130 height=72 align=left>
This section gives you easy access to your Sark Mail.  You can either open up a new web browser and access your 
<% if Session("UseMAPI") then %>
<a href="http://exchange.sarkcincinnati.com/exchange/LogonFrm.asp?mailbox=<%=session("username")%>" target="_blank">SARK email account</a>
<% else %>
SARK email account
<% end if %>
.<br>
Login with "<font color=navy><%=Session("username")%></font>".

<p>
<% if Session("UseMAPI") then %>
We have also added an e-mail viewer within the intranet.  Click 
<a href=email.asp>here</a>
to access this feature.
<% end if %>
<p>
Or, if you have an alternate mail account (i.e. Yahoo, Hotmail, etc.), you can access your mailbox
by logging into your account and using the information below to Pop your Sark mail.  If you have any questions
please use the feedback email form.
<hr size=1>

<table width="100%" cellpadding=0 cellspacing=5 border=0><tr><td valign=top><font size=1 face='ms sans serif, arial, geneva'>

<%emailSystem = request.cookies("emailSystem")%>
<!--
<form name="frmInfo" method="post" action="/IntranetDialogs/email/login.asp" target="Email" onSubmit="EmailPrep(); return true;">
<input type=hidden name=visit value="no">
<input type=hidden name=password value="">
<table border=0>
	<tr><td nowrap><font size=1 face='ms sans serif, arial, geneva'>Web Site:</font></td>
		<td>
		<select name="system">
			<option value="HOTMAIL"<%if emailSystem="HOTMAIL" then response.write(" selected")%>>Microsoft Hotmail</option>
			<option value="NETSCAPE"<%if emailSystem="NETSCAPE" then response.write(" selected")%>>Netscape Webmail</option>
		 	<option value="YAHOO"<%if emailSystem="YAHOO" then response.write(" selected")%>>Yahoo! Mail</option>
		</select>
		</td>
	</tr>
	<tr><td nowrap><font size=1 face='ms sans serif, arial, geneva'>Username:</font></td><td><input type=text name=username value="<%=request.cookies("emailUsername")%>" size=16 maxlength=16></td></tr>
	<tr><td nowrap><font size=1 face='ms sans serif, arial, geneva'>Password:</font></td><td><input type=password name=pass size=16 maxlength=16></td></tr>
	<tr><td>&nbsp;</td><td><input type=submit class=button value="Login"><input type=button class=button value="Visit" onClick="document.frmInfo.visit.value='yes'; document.frmInfo.submit(); document.frmInfo.visit.value='no'; return true;"></td></tr>
</table>
</form>
-->
</font></td><td width=5>&nbsp;</td><td valign=top>

<table class=tableShadow border=0 cellpadding=10 cellspacing=0><tr><td bgcolor=#ffffcc>
<font size=1 face='ms sans serif, arial, geneva'>
<font color=navy><b>To access mail from another system:</b></font><br>
Signup with the appropriate service (by visiting their site)
and use the following settings for POP3 access to your Sark mail:<p>
&nbsp;&nbsp;&#149;&nbsp;POP3: <b>exchange.sarkcincinnati.com</b><br>
&nbsp;&nbsp;&#149;&nbsp;POP User Name: <b><%=session("Username")%></b><br>
&nbsp;&nbsp;&#149;&nbsp;POP User Password: <b>[password]</b><br>
&nbsp;&nbsp;&#149;&nbsp;Port Number: <b>110</b><br>
</font>
</td></tr></table>

</td></tr></table>
<!--
<script language=javascript>
<!--
document.frmInfo.pass.focus()
// -->
</script>

<!-- #include file="../../footer.asp" -->
