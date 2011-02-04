<% response.buffer = true %>
<!-- #include file="../../script.asp" -->
<!-- #include file="CDOconstants.inc" -->
<!-- #include file="CDOlogin.asp" -->


<%
'---------------------------------------------------------------------------------
' Description:	Replies to, or forwards a specific message from a user's mailbox.
' History:		10/01/1999 - KDILL - Created
'				04/08/2000 - KDILL - Completed send form display and basic send functionality
'				04/11/2000 - KDILL - After sending it will return to the original msg
'				04/25/2000 - KDILL - Fixed SMTP lookup for internal people upon reply
'				05/02/2000 - KDILL - Fixed problem with circular replies
'				05/19/2000 - KDILL - Fixed problem with SMPT: showing up twice in replies
'---------------------------------------------------------------------------------



'---------------------------------------
'  Initialize objects					
'---------------------------------------
on error resume next
set objOMSession = Session("MAPIsession")
set objRenderApp = Application("RenderApplication")
set objRenderer = objRenderApp.CreateRenderer(2)
set objMessage = objOMSession.GetMessage(request("msg"))
objRenderer.Datasource = objMessage
set objFolder = objOMSession.GetFolder(objMessage.FolderID)
set objMessages = objFolder.Messages


'---------------------------------------
'  Extract addresses from original msg	
'---------------------------------------
addrTo = ""
addrCc = ""
cmd = Request("cmd")
if instr(cmd, "fwd") then Pre = "FW: "
strAddress = objMessage.Sender.Address
if InStr(strAddress, "cn=") then strAddress = objMessage.Sender.Fields(CdoPR_7BIT_DISPLAY_NAME_W) & "@sark.com"
if (instr(strAddress, "SMTP:") = 0) then strAddress = "SMTP:" & strAddress
addrFrom = objMessage.Sender.Name & " [" & strAddress & "]"
set objRecipients = objMessage.Recipients
for i = 1 to objRecipients.Count
	set objOneRecip = objRecipients.Item(i)
	strRecipient = objOneRecip.Name
	strAddress = objOneRecip.Address
	if instr(strAddress, "SMTP:") = 0 then strAddress = "SMTP:" & strAddress
	if InStr(strAddress, "cn=") then strAddress = objOneRecip.Fields(CdoPR_7BIT_DISPLAY_NAME_W) & "@sark.com"
	if InStr(lcase(strAddress), Session("username") & "@") = 0 then
		select case objOneRecip.Type
			case 1: addrTo = addrTo & "; " & strRecipient & " [" & strAddress & "]"
			case 2: addrCc = addrCc & "; " & strRecipient & " [" & strAddress & "]"
			end select
		end if
	next
addrTo = mid(addrTo, 3)
addrCc = mid(addrCc, 3)


'---------------------------------------
'  Decide on arrangement of addresses	
'---------------------------------------
title = "Send new message"
Pre = ""
newAddrTo = ""
newAddrCc = ""
newAddrBcc = ""
select case cmd
	case "fwd"
		Pre = "FW: "
		title = "Forward message"
	case "reply"
		Pre = "RE: "
		title = "Reply to message"
		newAddrTo = addrFrom
	case "replyall"
		Pre = "RE: "
		title = "Reply All to message"
		newAddrTo = addrFrom
		if addrTo <> "" then newAddrTo = newAddrTo & "; " & addrTo
		newAddrCc = addrCc
end select
if objMessage is not nothing then
	subject = objMessage.Subject
	if (instr(subject, pre) <> 1) then subject = pre & subject
	end if
%>


<html>
<head>
<title><% =title %></title>
<link rel="stylesheet" type="text/css" href="../../styles.htm">

<script language=javascript>
<!--
function init(){
<% if cmd = "" then %>
	document.frmInfo.to.focus()
<% else %>
	document.frmInfo.msgtext.focus()
<% end if %>
	}
// -->
</script>

</head>


<body topmargin=6 bgcolor=silver onLoad="init()">

<form name="frmInfo" method=post action="email_cmds.asp" onSubmit="return validate()">

<table class=tableShadow border=1 cellspacing=0 cellpadding=2 width="100%"><tr><td bgcolor=#e0e0e0 align=center>
<table border=0 width="100%">

	<tr>
		<td colspan=2>
			<table border=0 cellspacing=0 cellpadding=0 width="100%"><tr>
			<td align=left><input type=submit class=button value=" Send "></td>
			<td align=right><input type=button class=button value=" Cancel " onClick="window.history.back()">&nbsp;&nbsp;<input type=button class=button value=" Close " onClick="window.close()"></td>
			</tr></table>
		</td>
	</tr>

	<tr class="sectionBody">
		<td valign=top><b>From:</b></td>
		<td width="90%"><% =Session("Name") & " [SMTP:" & Session("Username") %>@sark.com]
			<input type=hidden name=from value="<% =Session("Username") %>@sark.com">
		</td>
	</tr>

	<tr class="sectionBody">
		<td valign=top><b>To:</b></td>
		<td width="90%"><input name=to class=edit value="<% =newAddrTo %>" size=70 maxlength=255>
			</font>
		</td>
	</tr>

	<tr class="sectionBody">
		<td valign=top><b>Cc:</b></td>
		<td><input name=cc class=edit value="<% =newAddrCc %>" size=70 maxlength=255></td>
	</tr>

	<tr class="sectionBody">
		<td><b>Subject:&nbsp;&nbsp;</b></td>
		<td><input name=subject class=edit value="<% =Subject%>" size=70 maxlength=255></td>
	</tr>

</table>
</td></tr></table>
<br>


<center>
<textarea name=msgtext class=edit rows=19 cols=73><%
if cmd <> "" then
	Response.write NL & NL & NL & "-----Original Message-----" & NL & "From: " & objMessage.Sender.Name & NL & "To: "
	objRenderer.RenderProperty CdoPR_DISPLAY_TO, 0, Response
	Response.write NL & "Sent: "
	objRenderer.RenderProperty CdoPR_CLIENT_SUBMIT_TIME, 0, Response
	Response.write NL & "Subject: " & objMessage.Subject & NL & NL & objMessage.Text
end if
%></textarea>
</center>

<input type=hidden name="msg" value="<%=request("msg")%>">
<input type=hidden name="unread" value="<%=request("unread")%>">
<input type=hidden name="cmd" value="send">
<input type=hidden name="subcmd" value="<%=cmd%>">
</form>

</body>
</html>

<script language=javascript>
<!--
function validate(){
	var result = true
	if (document.frmInfo.to.value == ""){
		alert("Please fill out the target email address.")
		document.frmInfo.to.focus()
		result = false
		}
	else if (document.frmInfo.subject.value == ""){
		alert("Please fill out the subject message.")
		document.frmInfo.subject.focus()
		result = false
		}
	return result
	}
// -->
</script>

<%
'---------------------------------------
'  Cleanup objects						
'---------------------------------------
on error resume next
objMessage.Unread = false
objMessage.Update
set objRenderApp = nothing
set objRenderer = nothing
set objFolder = nothing
set objMessage = nothing
set objMessages = nothing
set objTempMessage = nothing
%>