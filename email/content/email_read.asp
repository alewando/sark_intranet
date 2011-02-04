<% response.buffer = true %>
<!-- #include file="../../script.asp" -->
<!-- #include file="CDOconstants.inc" -->
<!-- #include file="CDOlogin.asp" -->


<%
'---------------------------------------------------------------------------------
' Description:	Displays a specific email message from the user's mailbox.
' History:		10/01/1999 - KDILL - Created
'				10/04/1999 - KDILL - Added Delete and Move functionality
'				04/06/2000 - KDILL - Rebuilt opener highlight to be optional
'				04/11/2000 - KDILL - Cleaned up implementation for Welcome unread section
'				04/26/2000 - KDILL - Added Toggle Header functionality
'---------------------------------------------------------------------------------



'---------------------------------------
'  Initialize objects					
'---------------------------------------
set objOMSession = Session("MAPIsession")
set objRenderApp = Application("RenderApplication")
set objRenderer = objRenderApp.CreateRenderer(2)
on error resume next
if Application("debug") then response.write NL & NL & "<!-- msg = " & request("msg") & " -->" & NL & NL
set objMessage = objOMSession.GetMessage(request("msg"))
if objMessage is nothing then response.redirect "email_delete_error.htm"
on error goto 0
objRenderer.Datasource = objMessage


'if session("username") = "kdill" then
'	response.write "<!-- Properties of msg: " & objMessage.Subject & NL
'	set objFields = objMessage.Fields
'	on error resume next
'	response.write "ConversationTopic: " & objMessage.ConversationTopic & NL
'	response.write "ConversationIndex: " & objMessage.ConversationIndex & NL
'	for i = 1 to objFields.count
'		response.write "Field " & i & ": " & objFields(i) & " (" & objFields(i).ID & ")" & NL
'		next
'	on error goto 0
'	response.write "-->" & NL & NL & NL
'	end if


'---------------------------------------
'  Determine prev / next messages		
'---------------------------------------
set objFolder = objOMSession.GetFolder(objMessage.FolderID)
set objMessages = objFolder.Messages
curmsg = 0
cont = true
msg_next = ""
msg_prev = ""
msg_next_unread = True
msg_prev_unread = True
unread = Request("unread")
displayUnread = (unread = "true")

if not displayUnread then
	' PREV / NEXT MSGS FOR EMAIL MAILBOX
	maxmsgs = objMessages.count
	while (curmsg < objMessages.count) and (cont)
		curmsg = curmsg + 1
		if objMessages.Item(curmsg).isSameAs(objMessage) then cont = false
		wend
	curRow = objMessages.count - curmsg
	if curmsg > 1 then
		msg_next = objMessages.Item(curmsg-1).ID
		msg_next_unread = objMessages.Item(curmsg-1).Unread
		end if
	if curmsg < objMessages.count then
		msg_prev = objMessages.Item(curmsg+1).ID
		msg_prev_unread = objMessages.Item(curmsg+1).Unread
		end if
else
	' PREV / NEXT MSGS FOR UNREAD EMAIL (WELCOME PAGE)
	maxmsgs = 0
	curmsg = 1
	for i = 1 to 15
		if Session("emailUnread" & i) <> "" then maxmsgs = maxmsgs + 1
		if Session("emailUnread" & i) = objMessage.ID then
			curmsg = maxmsgs
			actualLoc = i
			end if
		next
	curRow = curmsg - 1
	i = actualLoc + 1
	while (i < 15) and (Session("emailUnread" & i) = "")
		i = i + 1
		wend
	if (Session("emailUnread" & i) <> "") then msg_next = Session("emailUnread" & i)
	i = actualLoc - 1
	while (i > 1) and (Session("emailUnread" & i) = "")
		i = i - 1
		wend
	if (Session("emailUnread" & i) <> "") then msg_prev = Session("emailUnread" & i)
	end if


'---------------------------------------
'  Build "Move to" folder list			
'---------------------------------------
set objRootFolder = objOMSession.GetFolder(objOMSession.Inbox.FolderID)
item = 0
defaultfolder = 0
folderHtml = ""
sub DisplayFolderTree(folders, lvl)
	for each x in folders
		if x.class = 2 then
			item = item + 1
			folderHtml = folderHtml & "<option value='" & x.id & "'"
			if x.isSameAs(objFolder) then
				folderHtml = folderHtml & " selected"
				defaultfolder = item - 1
				end if
			folderHtml = folderHtml & ">"
			for i = 0 to lvl-1
				folderHtml = folderHtml & "&nbsp;&nbsp;"
				next
			folderHtml = folderHtml & x.Name & "</option>"
			if x.folders.count > 0 then
				DisplayFolderTree x.folders, lvl+1
				end if
			end if
		next
end sub
DisplayFolderTree objRootFolder.Folders, 0


title = "Message " & (curRow+1) & " of " & maxmsgs
if displayUnread then title = title & " unread msgs "
title = title & " in " & objFolder.Name
%>


<html>
<head>
<title><% =title %></title>
<link rel="stylesheet" type="text/css" href="../../styles.htm">

<script language=javascript>
<!--
var browserName = navigator.appName
var browserAbbr = "";
var browserVersion = parseInt(navigator.appVersion.substring(0,1))
if (browserName=="Microsoft Internet Explorer") {browserAbbr = "IE";}
else if (browserName=="Netscape") {browserAbbr = "NS";}
var dhtml = ((browserAbbr == "IE") && (browserVersion > 3))

var curRow = <% = curRow %>
var highlight = <% =lcase(Request("row") <> "") %>
if (highlight) window.opener.highlight = false
if (highlight) window.opener.rowHighlight(window.opener.document.all("row" + curRow))
function readMsg(msgid, row){
	if (highlight) window.opener.rowNormal(window.opener.document.all("row" + curRow))
	emailUrl = "email_read.asp?msg=" + msgid + "&unread=<% =unread %>"
	if (highlight) emailUrl += "&row=" + row
	window.location.href = emailUrl
	}
function terminateWin(){
	if (highlight){
		window.opener.highlight = true
		window.opener.rowNormal(window.opener.document.all('row' + curRow))
		}
	}
// -->
</script>

</head>


<body topmargin=6 bgcolor=silver onUnload="terminateWin()"><a name="top"></a>


<form name="frmInfo" method=post action="email_cmds.asp">

<table class=tableShadow border=0 cellspacing=0 cellpadding=2 width="100%"><tr><td bgcolor=#e0e0e0>
<table border=0 width="100%">

	<tr class="sectionBody">
		<td colspan=2>
			<table border=0 cellspacing=0 cellpadding=0 width="100%"><tr>
			<td align=left><input type=button class=button value=" Reply " onClick="window.location.href='email_send.asp?cmd=reply&unread=<%=unread%>&msg=<% =objMessage.ID %>'"><input type=button class=button value="Reply All" onClick="window.location.href='email_send.asp?cmd=replyall&unread=<%=unread%>&msg=<% =objMessage.ID %>'"><input type=button class=button value="Forward" onClick="window.location.href='email_send.asp?cmd=fwd&unread=<%=unread%>&msg=<% =objMessage.ID %>'"><input type=button class=button value="Delete" onClick="document.frmInfo.cmd.value='delete'; document.frmInfo.submit();"><% if Session("BrowserDHTML") then %><input type=button class=button value=" Print " onClick="window.print()"><% end if %></td>
			<td align=right><% if msg_prev <> "" then %><input type=button class=button value="  Prev  " onClick="readMsg('<% =msg_prev %>', <%=CInt(request("row"))-1 %>)"><% end if %><% if msg_next <> "" then %><input type=button class=button value="  Next  " onClick="readMsg('<% =msg_next %>', <%=CInt(request("row"))+1 %>)"><% end if %>&nbsp;&nbsp;<input type=button class=button value=" Close " onClick="window.close()"></td>
			</tr></table>
		</td>
	</tr>

	<tr class="sectionBody">
		<td valign=top><b>From:</b></td>
		<td width="90%">
<%
objRenderer.RenderProperty CdoPR_SENT_REPRESENTING_NAME_W, 0, Response
AddrType = objMessage.Fields(CdoPR_SENT_REPRESENTING_ADDRTYPE)
Address  = objMessage.Fields(CdoPR_SENT_REPRESENTING_EMAIL_ADDRESS)
If AddrType <> "" And AddrType <> "EX" And Address <> "" Then %>
&nbsp;&nbsp;[<% objRenderer.RenderProperty CdoPR_SENT_REPRESENTING_ADDRTYPE, 0, Response %>:<% objRenderer.RenderProperty CdoPR_SENT_REPRESENTING_EMAIL_ADDRESS, 0, Response %>]
<% End If %>
		</td>
	</tr>

	<tr class="sectionBody">
		<td valign=top><b>To:</b></td>
		<td><% objRenderer.RenderProperty CdoPR_DISPLAY_TO, 0, Response %></td>
	</tr>
<% if objMessage.Fields(CdoPR_DISPLAY_CC) <> "" then %>
	<tr class="sectionBody">
		<td valign=top><b>Cc:</b></td>
		<td><% objRenderer.RenderProperty CdoPR_DISPLAY_CC, 0, Response %></td>
	</tr>
<% end if %>
	<tr class="sectionBody">
		<td valign=top><b>Subject:&nbsp;&nbsp;</b></td>
		<td><% objRenderer.RenderProperty CdoPR_SUBJECT, 0, Response %></td>
	</tr>

	<tr class="sectionBody">
		<td valign=top><b>Sent:</b></td>
		<td><% objRenderer.RenderProperty CdoPR_CLIENT_SUBMIT_TIME, 0, Response %></td>
	</tr>

</table>
</td></tr></table>


<%
header = ""
on error resume next
header = objMessage.Fields(8192031)
if header <> "" then
	header = Replace(header, NL, "<br>")
	header = Replace(header, " with ", "<br>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;with ")
	header = Replace(header, "id ", "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;id ")
	end if
on error goto 0
%>

<table border=0 width="100%"><tr class="sectionBody">
<!--
	<td>Move message to:
		<select name="folderlist_footer" class=edit onChange="document.frmInfo.cmd.value='move'; document.frmInfo.destfolder.value=this.options[this.selectedIndex].value; document.frmInfo.submit();">
		<% =folderHtml %>
		</select>
		</td>
-->
<% if header <> "" then %>
	<td>[<a href="javascript: toggleHeader()">Toggle Header</a>]</td>
<% end if %>
	<td align=right>[<a href="#move">Bottom</a>]</td>
</tr></table>


<span id="header" style="display: none"><%=header%></span>

<center>
<%
strBody = objRenderer.RenderProperty( CdoPR_RTF_COMPRESSED, 0)
strBodyAttributes = objRenderer.RenderProperty( AMHTML_HTML_BODYTAG, 0)
nPos = CInt(InStr(strBodyAttributes, "bgcolor"))
response.write "<!-- nPos=" & nPos & " -->"
If strBodyAttributes = "" Or nPos = 0 Then strBodyAttributes = "bgcolor=white " & strBodyAttributes
%>
<table class=tableShadow border=0 cellpadding=5 cellspacing=0 width="100%"><tr class="sectionBody"><td <% =strBodyAttributes %> height=280 valign=top>
<% =strBody %>&nbsp;
</td></tr></table>
</center>


<table border=0 width="100%"><tr class="sectionBody">
	<td>Move message to:
		<select name="folderlist_footer" class=edit onChange="document.frmInfo.cmd.value='move'; document.frmInfo.destfolder.value=this.options[this.selectedIndex].value; document.frmInfo.submit();">
		<% =folderHtml %>
		</select>
		</td>
	<td align=right>[<a href="#top">Top</a>]</td>
</tr></table>


<a name="move"></a>

<input type=hidden name="msg_next_unread" value=<% =msg_next_unread %>>
<input type=hidden name="destfolder" value="">
<input type=hidden name="unread" value="<% =unread %>">
<input type=hidden name="cmd" value="">
<input type=hidden name="msg" value="<% =objMessage.ID %>">
<input type=hidden name="msg_next" value="<% =msg_next %>">

</form>

<script language=javascript>
<!--
function toggleHeader(){
	if (dhtml){
		if (header.style.display == ""){
			header.style.display = "none"
			}
		else{
			header.style.display = ""
			}
		}
	}
function resetFolderList(selectCtrl){
	selectCtrl.selectedIndex = <% =defaultfolder %>
	}
// -->
</script>

</body>
</html>

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
set objRootFolder = nothing
%>