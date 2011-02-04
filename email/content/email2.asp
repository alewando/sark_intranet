
<%
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
isDeleteFolder = (objDeleteFolder.isSameAs(objFolder))


emailError = (err > 0)
sectionTitle = "Mailbox, " & objFolder.Name
if emailError then sectionTitle = "Mailbox"



'---------------------------------------
'  Build folder listing					
'---------------------------------------
item = 0
folderHtml = ""
folderOptionList = ""
folderOptionListIndex = 0
curFolderLvl = 0
folderPath = ""
postSectionBtns = "<p><table class=tableShadow border=0 cellpadding=3 cellspacing=0 width=120><tr><td bgcolor=silver><b>Folders:</b> <span id='emailFolderCnt'></span></td></tr><tr><td bgcolor=white nowrap><span id=emailFolderList></span></td></tr></table><p>"
userFolder = (curFolderLvl > 0)
if not userFolder then
	if (objFolder.Name <> "Calendar") and (objFolder.Name <> "Contacts") and (objFolder.Name <> "Deleted Items") and (objFolder.Name <> "Inbox") and (objFolder.Name <> "Journal") and (objFolder.Name <> "Notes") and (objFolder.Name <> "Outbox") and (objFolder.Name <> "Sent Items") and (objFolder.Name <> "Tasks") then userFolder = true
	end if
%>

<!-- #include file="../section.asp" -->


<form name="frmInfo" method="post" action="email.asp">

<table class="tableShadow" border="0" cellpadding="2" cellspacing="0" width="100%"><tr bgcolor="#e0e0e0" height="25" class="sectionBody">

	<td width="20%" valign="center" nowrap>&nbsp;
		<a href="javascript: newMail()">
		<img src="images/mailbox/Newmail.gif" style="cursor:hand" alt="Compose new mail message" width="16" height="17" border="0" align="absmiddle">
		New msg</a>
		</td>

	<td width="20%" valign="center" nowrap>&nbsp;
		<a href="javascript: window.msgs.location.reload()">
		<img src="images/mailbox/Refresh2.gif" style="cursor:hand" alt="Refresh current folder message listing" width="13" height="16" border="0" align="absmiddle">
		Refresh</a>
		</td>

	<td width="20%" valign="center" nowrap>&nbsp;
		<a href="javascript: newFolder()">
		<img src="images/mailbox/Newfoldr.gif" style="cursor:hand" alt="Create a new subfolder" border="0" align="absmiddle" WIDTH="20" HEIGHT="20">
		New folder</a>
		</td>

	<td width="20%" valign="center" nowrap>&nbsp;
<% if userFolder then %>
		<a href="javascript: deleteFolder()">
		<img src="images/mailbox/Delfoldr.gif" style="cursor:hand" alt="Delete current folder and all messages" width="16" height="17" border="0" align="absmiddle">
		Delete folder</a>
<% end if %>
		</td>

	<td width="20%" valign="center" nowrap>&nbsp;
<% if userFolder then %>
		<a href="javascript: renameFolder()">
		<img src="images/mailbox/Delfoldr.gif" style="cursor:hand" alt="Rename current folder" width="16" height="17" border="0" align="absmiddle">
		Rename folder</a>
<% end if %>
		</td>

</tr></table>



<% if emailError then %>
<table border="0" width="100%"><tr height="350"><td><%=Application("EmailErrorMsg")%></td></tr></table>
<% else %>
<iframe name="msgs" src="email_msgs.asp?folder=<%=request("folder")%>" height="350" width="100%" frameborder="0" marginwidth="0" marginheight="0"></iframe>
<% end if %>



<table class="tableShadowSilver" border="0" width="100%" cellpadding="3" cellspacing="0" bgcolor="#e0e0e0"><tr height="20" class="sectionBody">
	<td width="100" nowrap>
		<span id="MoveFolder"></span>
		</td>
	<td nowrap>
		<a href="javascript:purgeEmail()"><img src="images/mailbox/urgent.gif" alt="PERMANENTLY delete the selected messages!" width="13" height="16" border="0" align="absmiddle"><% if not isDeleteFolder then %></a><a href="javascript:deleteEmail()"><img src="images/mailbox/DeleteMsg.gif" alt="Delete selected messages" height="14" width="14" border="0" align="absmiddle"><% end  if %>Delete</a>
		</td>
	<td nowrap>
		<a href="javascript:purgeFolder()"><img src="images/mailbox/empfldr.gif" alt="PERMANENTLY delete all messages from this folder!" height="20" width="20" border="0" align="absmiddle">Delete all</a>
		</td>
	<td width="90%" nowrap>&nbsp;</td>
	<td width="50" nowrap>
		<% if DisplayPage > 1 then %><a href="email.asp?folder=<% =request("folder") %>&amp;DisplayPage=<% =DisplayPage-1 %>">&lt;&lt;&nbsp;&nbsp;Prev</a><% end if %>
		</td>
	<td width="50" nowrap>
		<% if BaseTotal > BaseCountEnd then %><a href="email.asp?folder=<% =request("folder") %>&amp;DisplayPage=<% =DisplayPage+1 %>">Next&nbsp;&nbsp;&gt;&gt;</a><% end if %>
		</td>
</tr></table>




<!-- #include file="../../footer.asp"-->


<% if not emailError then %>
<script language="javascript">
<!--
function moveMsgs(ctrl){
	if ((ctrl.selectedIndex != <% =folderOptionListIndex %>) && (ctrl.selectedIndex > 0)){
		window.msgs.document.frmInfo.MoveFolder.value = ctrl.options[ctrl.selectedIndex].value
		window.msgs.document.frmInfo.cmd.value = "move"
		window.msgs.document.frmInfo.submit()
		ctrl.selectedIndex = 0
		}
	}

function deleteEmail(){
//	if (confirm("Move the selected messages to the 'Deleted Items' folder?")){
		window.msgs.document.frmInfo.cmd.value = "delete"
		window.msgs.document.frmInfo.submit()
//		}
	}

function purgeEmail(){
	if (confirm("PERMANENTLY delete the selected messages?")){
		window.msgs.document.frmInfo.cmd.value = "deleteperm"
		window.msgs.document.frmInfo.submit()
		}
	}

function purgeFolder(){
	if (confirm("PERMANENTLY delete all messages from this folder?")){
		window.msgs.document.frmInfo.cmd.value = "purgefolder"
		window.msgs.document.frmInfo.submit()
		}
	}

function newMail(){
	winTech= window.open("email_send.asp", "OpenEmail","height=450,width=580,toolbar=0,location=0,directories=0,status=1,menubar=0,resizable=1,scrollbars=1")
	}

function newFolder(){
	var folderName = prompt("Create new folder (place a ':' in front to create a root folder)...", "New folder")
	if ((folderName != null) && (folderName != "")){
		window.msgs.document.frmInfo.newfolder.value = folderName
		window.msgs.document.frmInfo.submit()
		}
	}

function deleteFolder(){
	if (confirm("Delete this folder and all messages in it?")){
		window.msgs.document.frmInfo.cmd.value = "deletefolder"
		window.msgs.document.frmInfo.submit()
		}
	}

function renameFolder(){
	var newName = prompt("Please enter the folder's new name.", "<% =objFolder.Name %>")
	if ((newName != null) && (newName != "<% =objFolder.Name %>")){
		window.msgs.document.frmInfo.renamefolder.value = newName
		window.msgs.document.frmInfo.submit()
		}
	}
	
// -->
</script>
<% end if %>


<%
'---------------------------------------
'  Cleanup objects						
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

