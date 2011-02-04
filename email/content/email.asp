<%
response.buffer = true
response.expires = 0
%>

<% if Session("UseMAPI") then %>

<!-- #include file="CDOconstants.inc" -->
<!-- #include file="CDOlogin.asp" -->


<%
'---------------------------------------------------------------------------------
' Description:	Displays the content of the user's sark mailbox.
' History:		09/30/1999 - KDILL - Created
'				11/05/1999 - KDILL - Revised interface; delete/move selected email(s)
'				11/10/1999 - KDILL - Fixed bug displaying messages in 2+ display page
'				                   - Added delete all messages
'---------------------------------------------------------------------------------



'---------------------------------------
'  Get start page to display			
'---------------------------------------
dim objDeleteFolder
RowsPerPage = 15
DisplayPage = 1
if request("DisplayPage") <> "" then DisplayPage = request("DisplayPage")



'---------------------------------------
'  Log on to MAPI session				
'---------------------------------------
set objOMSession = Session("MAPIsession")
if request("folder") = "" then
	set objFolder = Session("/Inbox")
else
	set objFolder = objOMSession.GetFolder(request("folder"))
end if
set objRootFolder = Session("/")
if isEmpty(Session("/Deleted Items")) then
	for each x in objRootFolder.Folders
		if x.Name = "Deleted Items" then set Session("/Deleted Items") = x
		next
	end if
set objDeleteFolder = Session("/Deleted Items")


'---------------------------------------
'  Delete emails from folder			
'---------------------------------------
if request("cmd") = "delete" then
	on error resume next
	for each x in request.form("chkmsgs")
		set objMessage = objOMSession.GetMessage(x)
		if err = 0 then
			objMessage.MoveTo(objDeleteFolder.id)
			end if
		next
	on error goto 0
	end if


'---------------------------------------
'  PERMANENLY delete emails from folder	
'---------------------------------------
if request("cmd") = "deleteperm" then
	on error resume next
	for each x in request.form("chkmsgs")
		set objMessage = objOMSession.GetMessage(x)
		if err = 0 then
			objMessage.Delete()
			set objMessage = nothing
			end if
		next
	on error goto 0
	end if


'-------------------------------------------
'  PERMANENLY purge all emails from folder	
'-------------------------------------------
if request("cmd") = "purgefolder" then
	on error resume next
	set objMessages = objFolder.Messages
	objMessages.Delete()
	set objMessages = nothing
	on error goto 0
	end if


'---------------------------------------
'  Move emails	to another folder		
'---------------------------------------
if request("cmd") = "move" then
	newFolderID = request.form("MoveFolder")
	on error resume next
	for each x in request.form("chkmsgs")
		set objMessage = objOMSession.GetMessage(x)
		if err = 0 then
			set objTempMessage = objMessage.MoveTo(newFolderID)
			end if
		next
	on error goto 0
	end if


'---------------------------------------
'  Create new folder					
'---------------------------------------
if request("cmd") = "deletefolder" then
	set objNewFolder = objOMSession.GetFolder(objFolder.FolderID)
	objFolder.Delete
	if objNewFolder.isSameAs(objRootFolder) then set objNewFolder = objOMSession.Inbox
	set objFolder = objNewFolder
	end if

'---------------------------------------
'  Create new folder					
'---------------------------------------
if request("renamefolder") <> "" then
	objFolder.Name = request("renamefolder")
	objFolder.Update
	end if

'---------------------------------------
'  Set messages datasource				
'---------------------------------------
set objMessages = objFolder.Messages
BaseTotal = objMessages.Count
BaseCount = (DisplayPage-1) * RowsPerPage
BaseCountEnd = BaseCount + RowsPerPage
if BaseCountEnd >  BaseTotal then BaseCountEnd = BaseTotal
if BaseTotal = 0 then
	BaseMsg = "none"
else
	BaseMsg = (BaseCount+1) & "-" & BaseCountEnd & " of " & BaseTotal
end if

sectionTitle = "Mailbox, " & objFolder.Name & "&nbsp;&nbsp;<font size=1 face='ms sans serif, arial, geneva'>(" & BaseMsg & ")</font>"
%>
<!-- #include file="../section.asp" -->
<%
'---------------------------------------
'  Build folder listing					
'---------------------------------------
item = 0
folderHtml = ""
folderOptionList = ""
folderOptionListIndex = 0
curFolderLvl = 0
folderPath = ""
sub DisplayFolderTree(folders, lvl, path)
	for each x in folders
		if (lvl=0) and (x.name="Deleted Items") then
			set objDeleteFolder = x
			end if
		item = item + 1
		folderOptionList = folderOptionList & "<option value='" & x.ID & "'>"
		for i = 0 to lvl-1
			folderHtml = folderHtml & "&nbsp;&nbsp;&nbsp;"
			folderOptionList = folderOptionList & "&nbsp;&nbsp;"
			next
		folderOptionList = folderOptionList & x.Name & "</option>"
		itemHtml = "" & x.Name
		newPath = path & "/" & x.Name
		if folderName <> "" then
			if x.isSameAs(objNewFolder) then itemHtml = "<span style='color:red'>" & itemHtml & "</span>"
			end if
		if x.isSameAs(objFolder) then
			folderOptionListIndex = item
			folderHtml = folderHtml & "<b>" & itemHtml & "</b>"
			curFolderLvl = lvl
			folderPath = newPath
		else
			folderHtml = folderHtml & "<a href='email.asp?folder=" & x.id & "'>" & itemHtml & "</a>"
		end if
		folderHtml = folderHtml & "&nbsp;&nbsp;<font color=silver>" & x.messages.count & "</font><br>" & NL
		if x.folders.count > 0 then
			DisplayFolderTree x.folders, lvl+1, newPath
			end if
		next
end sub
DisplayFolderTree objRootFolder.Folders, 0, ""
response.write("<!-- FOLDER PATH = '" & folderPath & "' -->")
folderHtml = "<table class=tableShadow border=0 cellpadding=3 cellspacing=0 width=120><tr><td bgcolor=silver><font size=1 face='ms sans serif, arial, geneva'><b>Folders:</b> " & item & "</td></tr><tr><td bgcolor=white nowrap><font size=1 face='ms sans serif, arial, geneva'>" & NL & folderHtml
folderHtml = folderHtml & "</font></td></tr></table><p>" & NL
sectionBtns = sectionBtns & "<p>" & folderHtml
userFolder = (curFolderLvl > 0)
if not userFolder then
	if (objFolder.Name <> "Calendar") and (objFolder.Name <> "Contacts") and (objFolder.Name <> "Deleted Items") and (objFolder.Name <> "Inbox") and (objFolder.Name <> "Journal") and (objFolder.Name <> "Notes") and (objFolder.Name <> "Outbox") and (objFolder.Name <> "Sent Items") and (objFolder.Name <> "Tasks") then userFolder = true
	end if


'---------------------------------------
'  Create new folder					
'---------------------------------------
on error resume next
folderName = request("newfolder")
if folderName <> "" then
	on error resume next
	if left(folderName,1) = ":" then
		folderName = mid(folderName,2)
		set objNewFolder = objRootFolder.Folders.Add(folderName)
	else
		set objNewFolder = objFolder.Folders.Add(folderName)
	end if
	objNewFolder.Update
	if err > 0 then folderName = ""
	on error goto 0
	end if
%>

<script language=javascript>
<!--

var curRowObj = null
var highlight = true
function rowHighlight(rowObj){
	rowObj.style.background = "#ffffcc"
	curRowObj = rowObj
	}
function rowNormal(rowObj){
	if (rowObj == null) rowObj = curRowObj
	rowObj.style.background = "white"
	}


function displayMsg(row, obj){
	highlight = false
	winTech= window.open("email_read.asp?msg=" + obj + "&row=" + row, "OpenEmail","height=450,width=520,toolbar=0,location=0,directories=0,status=1,menubar=0,resizable=1,scrollbars=1")
	}

function writeDateTime(val){
	i = val.indexOf("/")
	emailDate = ""
	if (i > 0){
		if (i == 1) emailDate = "&nbsp;"
		emailDate = val.substring(0, i+1)
		val = val.substring(i+1, 100)
		}
	i = val.indexOf("/")
	if (i > 0){
		if (i == 1) emailDate += "0"
		emailDate += val.substring(0, i)
		val = val.substring(i+1, 100)
		}
	i = val.indexOf(" ")
	if (i > 0) val = val.substring(i+1, 100)
	i = val.indexOf(" ")
	if (i > 0){
		if (val.indexOf(":") == 1) emailDate += "&nbsp;&nbsp;"
		emailDate += "&nbsp;&nbsp;" + val.substring(0, i)
		if (val.substring(i,i+1) == "A") {emailDate += "a"}
		else {emailDate += "p"}
		}
	document.write(emailDate)
	}

var toggleon = true
function toggleMsgs(){
	for (i=0; i < document.frmInfo.elements.length; i++){
		x = document.frmInfo.elements[i]
		if (x.name == "chkmsgs"){
			x.checked = toggleon
			}
		}
	if (toggleon) {toggleon = false}
	else {toggleon = true}
	}

function displayRowMsg(rownum){
	var objval = document.all("obj" + rownum).value
	displayMsg(rownum, objval)
	}

var rowCnt = <% =((DisplayPage - 1) * RowsPerPage) %>
function displayRow(){
	document.write("<tr id=row" + rowCnt + " bgcolor=FFFFFF valign=top align=left onMouseOver='if (highlight) rowHighlight(this)' onMouseOut='if (highlight) rowNormal(this)' onClick='displayRowMsg(" + rowCnt + ")' style='cursor:hand'>")
	rowCnt++
	}

// -->
</script>
<form name="frmInfo" method=post action="email.asp">
<table class=tableShadow border=0 cellpadding=2 cellspacing=0 width="100%">
<font size=-1><b>Click <font color=gray size=-3>V</font> to select all.</b></font>
	<tr bgcolor=#e0e0e0 height=20>
	<td width="20%" valign=center nowrap><font size=1 face="ms sans serif, arial, geneva">&nbsp;
		<a href="javascript: newMail()">
		<img src="images/mailbox/Newmail.gif" style="cursor:hand" alt="Compose new mail message" width=16 height=17 border=0 align=absmiddle>
		New msg</a>
		</font></td>

	<td width="20%" valign=center nowrap><font size=1 face="ms sans serif, arial, geneva">&nbsp;
		<a href="javascript: window.location.href='email.asp?folder=<% =objFolder.ID %>&DisplayPage=<% =DisplayPage %>'">
		<img src="images/mailbox/Refresh2.gif" style="cursor:hand" alt="Refresh current folder message listing" width=13 height=16 border=0 align=absmiddle>
		Refresh</a>
		</font></td>

	<td width="20%" valign=center nowrap><font size=1 face="ms sans serif, arial, geneva">&nbsp;
		<a href="javascript: newFolder()">
		<img src="images/mailbox/Newfoldr.gif" style="cursor:hand" alt="Create a new subfolder" width=16 height=17 border=0 align=absmiddle>
		New folder</a>
		</font></td>

	<td width="20%" valign=center nowrap><font size=1 face="ms sans serif, arial, geneva">&nbsp;
<% if userFolder then %>
		<a href="javascript: deleteFolder()">
		<img src="images/mailbox/Delfoldr.gif" style="cursor:hand" alt="Delete current folder and all messages" width=16 height=17 border=0 align=absmiddle>
		Delete folder</a>
<% end if %>
		</font></td>

	<td width="20%" valign=center nowrap><font size=1 face="ms sans serif, arial, geneva">&nbsp;
<% if userFolder then %>
		<a href="javascript: renameFolder()">
		<img src="images/mailbox/Delfoldr.gif" style="cursor:hand" alt="Rename current folder" width=16 height=17 border=0 align=absmiddle>
		Rename folder</a>
<% end if %>
		</font></td>

</tr></table>



<%
'---------------------------------------
'  Display message count for folder		
'---------------------------------------
%>
<!--
<table border=0 width="100%" cellpadding=5 cellspacing=0><tr>
	<td><font size=1>
<% if DisplayPage > 1 then %><a href="email.asp?folder=<% =request("folder") %>&DisplayPage=<% =DisplayPage-1 %>">&lt;&lt;&nbsp;&nbsp;Prev</a><% end if %>
		</font></td>
	<td align=center><font size=1>&nbsp;
		</font></td>
	<td align=right><font size=1>
<% if BaseTotal > BaseCountEnd then %><a href="email.asp?folder=<% =request("folder") %>&DisplayPage=<% =DisplayPage+1 %>">Next&nbsp;&nbsp;&gt;&gt;</a><% end if %>
		</font></td>
</tr></table>
-->


<%
if err = 0 then


if (objFolder.Name = "Calendar") and (curFolderLvl = 0) then
	
	' INSERT CODE to render calendar folder!
	response.write "RENDER CALENDAR!"
	
else

'---------------------------------------
'  Create rendering objects				
'---------------------------------------
set objRenderApp = Application("RenderApplication")
set objRenderer = objRenderApp.CreateRenderer(3)


'---------------------------------------
'  Set icons for message priorities		
'---------------------------------------
set objFormat = objRenderer.Formats.Add(CdoPR_IMPORTANCE, Null)
set objPatterns = objFormat.Patterns
objPatterns.Add CdoLow, "<img src='images/mailbox/low.gif' alt='Low priority' width=13 height=16 border=0>"
objPatterns.Add CdoNormal, ""
objPatterns.Add CdoHigh, "<img src='images/mailbox/urgent.gif' alt='URGENT!' width=13 height=16 border=0>"


'---------------------------------------
'  Set icons for message attachments	
'---------------------------------------
set objFormat = objRenderer.Formats.Add(CdoPR_HASATTACH, Null)
set objPatterns = objFormat.Patterns
objPatterns.Add 0, ""
objPatterns.Add 1, "<img src='images/mailbox/papclip.gif' alt='Attachment(s)' width=10 height=13 border=0>"


'---------------------------------------
'  Set icons for message type			
'---------------------------------------
set objFormat = objRenderer.Formats.Add(CdoPR_MESSAGE_CLASS, Null)
set objPatterns = objFormat.Patterns
objPatterns.Add "IPM.Note",								"<img id=note%rowid% src='images/mailbox/envelope.gif' width=16 height=16 border=0>"
objPatterns.Add "IPM.Note.Secure",						"<img id=note%rowid% src='images/mailbox/encrypt.gif' width=16 height=16 border=0>"
objPatterns.Add "IPM.Note.Secure.Sign",					"<img id=note%rowid% src='images/mailbox/digi_sig.gif' width=16 height=16 border=0>"
objPatterns.Add "IPM.Note.Rules.OofTemplate.Microsoft",	"<img id=note%rowid% src='images/mailbox/oof.gif' width=16 height=16 border=0>"
objPatterns.Add "IPM.Post",								"<img id=note%rowid% src='images/mailbox/post.gif' width=16 height=16 border=0>"
objPatterns.Add "IPM.Document.*",						"<img id=note%rowid% src='images/mailbox/freedoc.gif' width=16 height=16 border=0>"
objPatterns.Add "IPM.Schedule.Meeting.Request",			"<img id=note%rowid% src='images/mailbox/mtgreq.gif' width=16 height=16 border=0>"
objPatterns.Add "IPM.Schedule.Meeting.Canceled",		"<img id=note%rowid% src='images/mailbox/mtgcancl.gif' width=16 height=16 border=0>"
objPatterns.Add "IPM.Schedule.Meeting.Resp.Neg",		"<img id=note%rowid% src='images/mailbox/mtgdecln.gif' width=16 height=16 border=0>"
objPatterns.Add "IPM.Schedule.Meeting.Resp.*",			"<img id=note%rowid% src='images/mailbox/mtgaccpt.gif' width=16 height=16 border=0>"
objPatterns.Add "*",									"<img id=note%rowid% src='images/mailbox/envelope.gif' width=16 height=16 border=0>"


'---------------------------------------
'  Set header icons and look and feel	
'---------------------------------------
'item_template = "<font size=1 face='ms sans serif, arial, geneva'>" & chr(34) & "displayMsg(%rowid%, '%obj%')" & chr(34) & " style='cursor:hand'>%value%</span></font>"
item_template = "<font size=1 face='ms sans serif, arial, geneva'>%value%</font>"
objRenderer.Datasource = objMessages
set objView = objRenderer.Views.Item(1)
objView.Columns.Item(1).Name = "<img src=images/mailbox/urgent.gif alt='Priority' width=13 height=16 border=0>"
objView.Columns.Item(2).Name = "<img src=images/mailbox/envelope.gif alt='Msg (file size)' width=16 height=16 border=0>"
objView.Columns.Item(3).Name = "<img src=images/mailbox/papclip.gif alt='Attachment(s)' width=10 height=13 border=0>"
'objView.Columns.Item(4).Name = "<img src=images/mailbox/meeting.gif width=12 height=13 border=0>"
objView.Columns.Item(4).RenderUsing = item_template
objView.Columns.Item(4).Width = 15
objView.Columns.Item(5).RenderUsing = item_template
objView.Columns.Item(6).RenderUsing = "<font size=1 face='ms sans serif, arial, geneva'><span onClick=" & chr(34) & "displayMsg(%rowid%, '%obj%')" & chr(34) & " style='cursor:hand'><script>writeDateTime('%value%')</script></span></font>"
objView.Columns.Item(6).Width = 8
objView.Columns.Item(7).RenderUsing = "<script>document.all('note%rowid%').alt = '%value% bytes'</script>"
objView.Columns.Item(7).Name = ""
objView.Columns.Item(7).Width = 0


'---------------------------------------
'  Create checkbox column				
'---------------------------------------
set newCol = objView.Columns.Add("&nbsp;<span onClick='javascript:toggleMsgs()' style='color:gray; cursor:hand'>&nbsp;V&nbsp;</span>", CdoPR_MESSAGE_CLASS, 10, 8, 0)
newCol.RenderUsing = "<input type=checkbox name=chkmsgs value='%obj%' style='cursor:default' onClick='toggleon=true; self.event.cancelBubble=true'><input type=hidden name=obj%rowid% value='%obj%'>"


'---------------------------------------
'  Set general layout, and render!		
'---------------------------------------
objRenderer.TablePrefix = "<table class=tableShadow border=0 cellpadding=0 height=380 cellspacing=0 style='border-bottom:0px; border-right:1px silver solid'><tr><td valign=top><table border=0 cellpadding=1 cellspacing=0>"
objRenderer.TableSuffix = "</table></td></tr></table>"
objRenderer.HeadingRowPrefix = "<tr bgcolor=#f0f0f0>"
objRenderer.HeadingCellPattern = "<font size=1 face='ms sans serif, arial, geneva'>%value%</font>"
objRenderer.RowPrefix = "<script>displayRow()</script>"
objRenderer.RowsPerPage = RowsPerPage
objRenderer.Render 1, DisplayPage, 0, Response
%>

<table class=tableShadow border=0 width="100%" cellpadding=3 cellspacing=0 bgcolor="#e0e0e0"><tr height=20>
	<td nowrap>
		<select name=MoveFolder class=edit onChange="moveMsgs(this)">
		<option value="">Move messages to...</option>
		<% =folderOptionList %>
		</select>
		</td>
	<td nowrap><font size=1 face='ms sans serif, arial, geneva'>
	<%if objFolder.Name = "Deleted Items" then%>
		<a href="javascript:purgeEmail()"><img src="images/mailbox/DeleteMsg.gif" alt="PERMANENTLY delete the selected messages!" width=13 height=16 border=0 align=absmiddle>Delete</a><font color=blue> -permanently delete selected messages</font>
	<%else%>
		<a href="javascript:deleteEmail()"><img src="images/mailbox/DeleteMsg.gif" alt="Delete selected messages" height=14 width=14 border=0 align=absmiddle>Delete</a><font color=blue> -move selected messages to "Deleted Items"</font>
	<%end if%>
		</font></td>
	<!--<td nowrap><font size=1 face='ms sans serif, arial, geneva'>
		<a href="javascript:purgeFolder()"><img src="images/mailbox/empfldr.gif" alt="PERMANENTLY delete all messages from this folder!" height=20 width=20 border=0 align=absmiddle>Delete all</a>
		</font></td>
	-->
	<td width="90%" nowrap>&nbsp;</td>
	<td width=50 nowrap><font size=1 face='ms sans serif, arial, geneva'>
		<% if DisplayPage > 1 then %><a href="email.asp?folder=<% =request("folder") %>&DisplayPage=<% =DisplayPage-1 %>">&lt;&lt;&nbsp;&nbsp;Prev</a><% end if %>
		<% if BaseTotal > BaseCountEnd then %><a href="email.asp?folder=<% =request("folder") %>&DisplayPage=<% =DisplayPage+1 %>">Next&nbsp;&nbsp;&gt;&gt;</a><% end if %>
	</font></td>
</tr></table>
<%
end if

end if
%>

<input type=hidden name="newfolder" value="">
<input type=hidden name="renamefolder" value="">
<input type=hidden name="cmd" value="">
<input type=hidden name="folder" value="<% =objFolder.ID %>">
<input type=hidden name="msgcount" value="<% =objMessages.count %>">
</form>


<!-- #include file="../../footer.asp" -->

<script language=javascript>
<!--
function moveMsgs(ctrl){
	if ((ctrl.selectedIndex != <% =folderOptionListIndex %>) && (ctrl.selectedIndex > 0)){
		document.frmInfo.cmd.value = "move"
		document.frmInfo.submit()
		}
	else {alert("Please select a different folder.")}
	}

function deleteEmail(){
	if (confirm("Move the selected messages to the 'Deleted Items' folder?")){
		document.frmInfo.cmd.value = "delete"
		document.frmInfo.submit()
		}
	}

function purgeEmail(){
	if (confirm("PERMANENTLY delete the selected messages?")){
		document.frmInfo.cmd.value = "deleteperm"
		document.frmInfo.submit()
		}
	}

function purgeFolder(){
	if (confirm("PERMANENTLY delete all messages from this folder?")){
		document.frmInfo.cmd.value = "purgefolder"
		document.frmInfo.submit()
		}
	}

function newMail(){
	winTech= window.open("email_send.asp", "OpenEmail","height=500,width=640,toolbar=0,location=0,directories=0,status=1,menubar=0,resizable=1,scrollbars=1")
	}

function newFolder(){
	var folderName = prompt("Create new folder (place a ':' in front to create a root folder)...", "New folder")
	if ((folderName != null) && (folderName != "")){
		document.frmInfo.newfolder.value = folderName
		document.frmInfo.submit()
		}
	}

function deleteFolder(){
	if (confirm("Delete this folder and all messages in it?")){
		document.frmInfo.cmd.value = "deletefolder"
		document.frmInfo.submit()
		}
	}

function renameFolder(){
	var newName = prompt("Please enter the folder's new name.", "<% =objFolder.Name %>")
	if ((newName != null) && (newName != "<% =objFolder.Name %>")){
		document.frmInfo.renamefolder.value = newName
		document.frmInfo.submit()
		}
	}
// -->
</script>

<%
'---------------------------------------
'  Cleanup objects and logoff			
'---------------------------------------
on error resume next
set objRenderer = nothing
set objRenderApp = nothing
set objMessage = nothing
set objRootFolder = nothing
set objDeleteFolder = nothing
set objMessages = nothing
set objFolder = nothing
set objNewFolder = nothing
%>


<% else 'else clause of If Session("UseMAPI") %>
Unable to connect to mailbox
<% end if %>