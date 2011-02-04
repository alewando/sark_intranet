<% response.buffer = true %>
<!-- #include file="../../script.asp" -->
<!-- #include file="CDOconstants.inc" -->
<!-- #include file="CDOlogin.asp" -->


<%
'---------------------------------------------------------------------------------
' Description:	Executes a command on an email message from the user's mailbox.
' History:		10/01/1999 - KDILL - Created;  added Delete and Move functionality
'				04/11/2000 - KDILL - Cleaned up implementation for Welcome unread section
'				04/26/2000 - KDILL - Will now properly Forward attachments
'				04/28/2000 - KDILL - Sender information now displayed in message
'---------------------------------------------------------------------------------



'---------------------------------------
'  Initialize objects					
'---------------------------------------
on error resume next
set objOMSession = Session("MAPIsession")
on error goto 0
msg = request("msg")
cmd = request("cmd")
subcmd = request("subcmd")
%>


<html>
<head>
<script language=javascript>
<!--
function display(){
<% if (cmd="send") and (request("unread") <> "") then %>
	window.location.href="email_read.asp?msg=<% =msg %>&unread=<% =request("unread") %>"
<% else %>
	<% if (request("msg_next") <> "") and (request("msg_next_unread") = "True") then %>
		window.location.href="email_read.asp?msg=<% =request("msg_next") %>&unread=<% =request("unread") %>"
	<% else %>
		window.opener.location.href = window.opener.location.href
		i = setTimeout("window.close()", 700)
	<% end if %>
<% end if %>
	}
// -->
</script>

</head>

<body bgcolor=silver onLoad="display()">

<center><font size=4 face='ms sans serif, arial, geneva'>
<%
'---------------------------------------
'  Send email message					
'---------------------------------------
if cmd = "send" then
	response.write "Sending message..."
	if request("msg") <> "" then set objOriginalMessage = objOMSession.GetMessage(request("msg"))
	select case (request("subcmd"))
		case "fwd":	set objMessage = objOriginalMessage.Forward()
		case else:	set objMessage = Session("/Inbox").Messages.Add
	end select
	objMessage.Sender = objOMSession.CurrentUser
	objMessage.Subject = Request.Form("subject")
	objMessage.Text = Request.Form("msgtext")
	objMessage.Recipients.Delete
	objMessage.Recipients.AddMultiple Request.Form("to"), 1
	objMessage.Recipients.AddMultiple Request.Form("cc"), 2
	objMessage.Recipients.AddMultiple Request.Form("bcc"), 3
	objMessage.Recipients.Resolve
	objMessage.Update
	objMessage.Send
	set objOriginalMessage = nothing
else

	set objMessage = objOMSession.GetMessage(request("msg"))
	
	'---------------------------------------
	'  Move message to another folder		
	'---------------------------------------
	if cmd = "move" then
		response.write "Moving message..."
		set objNewMessage = objMessage.MoveTo(request("destfolder"))
		end if
	
	'---------------------------------------
	'  Delete email message					
	'---------------------------------------
	if cmd = "delete" then
		response.write "Deleting message..."
		'objMessage.Delete	'PERMANENT delete
		for i = 1 to 15
			if Session("emailUnread" & i) = objMessage.ID then
				Session("emailUnread" & i) = ""
			else
				lastValid = i
				end if
			next
		Session("emailUnreadCnt") = lastValid
		objMessage.MoveTo(Session("/Deleted Items").ID)
	end if

end if


on error resume next
set objMessage = nothing
set objNewMessage = nothing
response.flush
%>

</font></center>

</body>
</html>
