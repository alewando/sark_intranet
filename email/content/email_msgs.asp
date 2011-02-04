<%
'---------------------------------------------------------------------------------
' Description:	Displays the content of the user's sark mailbox.
' History:		09/30/1999 - KDILL - Created
'				11/05/1999 - KDILL - Revised interface; delete/move selected email(s)
'				11/10/1999 - KDILL - Fixed bug displaying messages in 2+ display page
'				                   - Added delete all messages
'				04/07/2000 - KDILL - Fixed timestamp display for am/pm (all displayed as pm)
'				05/04/2000 - KDILL - Added check for Exchange mail server down
'---------------------------------------------------------------------------------
%>

<!-- #include file="../../script.asp"-->
<% DebugHeader %>



<!-- #include file="CDOconstants.inc" -->
<!-- #include file="CDOlogin.asp" -->
<%
'---------------------------------------
'  Get start page to display			
'---------------------------------------
dim objDeleteFolder
RowsPerPage = 200
DisplayPage = 1
cmd = Request("cmd")
if request("DisplayPage") <> "" then DisplayPage = request("DisplayPage")


'---------------------------------------
'  Log on to MAPI session				
'---------------------------------------
on error resume next
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
if cmd = "delete" then
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
if cmd = "deleteperm" then
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
if cmd = "purgefolder" then
	on error resume next
	set objMessages = objFolder.Messages
	objMessages.Delete()
	set objMessages = nothing
	on error goto 0
	end if


'---------------------------------------
'  Move emails	to another folder		
'---------------------------------------
if cmd = "move" then
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
if cmd = "deletefolder" then
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
'  Create new folder					
'---------------------------------------
'on error resume next
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
			folderHtml = folderHtml & "<a href='email2.asp?folder=" & x.id & "'>" & itemHtml & "</a>"
		end if
		folderHtml = folderHtml & "&nbsp;&nbsp;<font color=silver>" & x.messages.count & "</font><br>"
		if x.folders.count > 0 then
			DisplayFolderTree x.folders, lvl+1, newPath
			end if
		next
end sub
DisplayFolderTree objRootFolder.Folders, 0, ""
%>

<html>
<head>
<!-- #include file="../../styles.htm" -->

<script language="javascript">
<!--

parent.emailFolderCnt.innerHTML = "<% =item %>"
parent.emailFolderList.innerHTML = "<% =folderHtml%>"
parent.MoveFolder.innerHTML = "<select name=MoveFolder class=edit onChange='moveMsgs(this)'><option value=''>Move messages to...</option><% =folderOptionList %></select>"

var curRowObj = null
var highlight = true
function rowHighlight(rowObj){
	if (rowObj != null){
		rowObj.style.background = "#ffffcc"
		curRowObj = rowObj
		}
	}
function rowNormal(rowObj){
	if (rowObj != null){
		if (rowObj == null) rowObj = curRowObj
		rowObj.style.background = "white"
		}
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
	i = val.lastIndexOf(" ")
	if (i > 0){
		if (val.indexOf(":") == 1) emailDate += "&nbsp;&nbsp;"
		emailDate += "&nbsp;" + val.substring(0, i)
		if (val.substring(i+1,i+2) == "A") {emailDate += "a"}
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
	document.write("<tr class=sectionBody id=row" + rowCnt + " bgcolor=FFFFFF valign=top align=left onMouseOver='if (highlight) rowHighlight(this)' onMouseOut='if (highlight) rowNormal(this)' onClick='displayRowMsg(" + rowCnt + ")' style='cursor:hand'>")
	rowCnt++
	}

// -->
</script>

</head>


<body bgcolor=#f0f0f0>

<center><span class="hide" style="font-family: arial; 8pt;">Loading folder messages...</span></center>


<form name=frmInfo method=post action="">

<%
response.flush
'---------------------------------------
'  Create rendering objects				
'---------------------------------------
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
objRenderer.Datasource = objMessages
set objView = objRenderer.Views.Item(1)
objView.Columns.Item(1).Name = "<img src=images/mailbox/urgent.gif alt='Priority' width=13 height=16 border=0>"
objView.Columns.Item(2).Name = "<img src=images/mailbox/envelope.gif alt='Msg (file size)' width=16 height=16 border=0>"
objView.Columns.Item(3).Name = "<img src=images/mailbox/papclip.gif alt='Attachment(s)' width=10 height=13 border=0>"
objView.Columns.Item(4).Width = 15
objView.Columns.Item(6).RenderUsing = "<span onClick=" & chr(34) & "displayMsg(%rowid%, '%obj%')" & chr(34) & " style='cursor:hand'><script>writeDateTime('%value%')</script></span>"
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
objRenderer.TablePrefix = "<table class=tableShadowWhite border=0 cellpadding=0 height='100%' cellspacing=0 style='border-bottom:0px; border-right:1px silver solid'><tr><td valign=top><table border=0 cellpadding=1 cellspacing=0>"
objRenderer.TableSuffix = "</table></td></tr></table>"
objRenderer.HeadingRowPrefix = "<tr bgcolor=#f0f0f0 class=sectionBody>"
objRenderer.HeadingCellPattern = "%value%"
objRenderer.RowPrefix = "<script>displayRow()</script>"
objRenderer.RowsPerPage = RowsPerPage
objRenderer.Render 1, DisplayPage, 0, Response
%>

<input type="hidden" name="MoveFolder" value="">
<input type="hidden" name="newfolder" value="">
<input type="hidden" name="renamefolder" value="">
<input type="hidden" name="cmd" value="">
<input type="hidden" name="folder" value="<% =objFolder.ID %>">
<input type="hidden" name="msgcount" value="<% =objMessages.count %>">

</form>
</body>
</html>

<style>
<!--
.hide {display: none}
-->
</style>

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

