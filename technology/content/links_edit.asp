<!--
Developer:    SSMITH
Date:         12/03/1998
Description:  Edit/Delete link to tech expert page
-->

<!-- #include file="../../script.asp" -->

<html>

<head>
<title>Edit Technology Link</title>
<!-- #include file="../../style.htm" -->
</head>


<% If Request("State") = "" Then %>
<body bgcolor=silver><center>
<script language=javascript>
<!--
function checkForm(){
	with (document.frmInfo){
		if (URL.value == "" || URL.value == null || URL.value == "http://") {
			alert("You must enter a value for URL");
			URL.focus();
			return false;
			}
		if (Title.value == "" || Title.value == null) {
			alert("You must enter a value for Title.");
			Title.focus();
			return false;
			}
		if (Description.value == "" || Description.value == null) {
			alert("You must enter a value for Description.");
			Description.focus();
			return false;
			}
		if (URL.value.indexOf("<") > 0 || URL.value.indexOf(">") > 0) {
			alert("You may not use HTML tags within the URL.");
			URL.focus();
			return false;
			}
		if (Title.value.indexOf("<") > 0 || Title.value.indexOf(">") > 0) {
			alert("You may not use HTML tags within the Title.");
			Title.focus();
			return false;
			}
		if (Description.value.indexOf("<") > 0 || Description.value.indexOf(">") > 0) {
			alert("You may not use HTML tags within the Description.");
			Description.focus();
			return false;
			}
		}
	document.frmInfo.submit();
	return true;
	}

function DeleteLink(){
	if (confirm("If you delete the link, an automated email message will be sent to the experts informing them of the deletion.  Is it OK to go ahead and delete the link?")) {
		var reason = prompt("Why are you deleting this link?", "")
		if (reason != null){
			if (reason == "") {alert("You must enter a reason for deleting this link!")}
			else{
				if (confirm("The system will automatically notify the 'experts' and alert them to the deletion of this link.")){
					document.frmInfo.Reason.value = reason
					document.frmInfo.Delete.value = 'true'
					document.frmInfo.submit()
					}
				}
			}
		}
	}
// -->
</script>

<%
sql = "Select * FROM Tech_Links WHERE Tech_Links_ID = " & Request("Link_ID")
set rs = DBQuery(sql)
%>

<table border=0 width='90%'><tr><td><font size=1 face="ms sans serif, arial, geneva"><b>
Link Information:
</b></font></td></tr></table> 

<form NAME="frmInfo" method="Post" ACTION="links_edit.asp">

<INPUT TYPE="HIDDEN" NAME="EmployeeName" VALUE="<%=Request("empname")%>">
<INPUT TYPE="HIDDEN" NAME="Tech_ID" VALUE="<%=Request("Tech_ID")%>">
<INPUT TYPE="HIDDEN" NAME="State" VALUE="Update"> 
<INPUT TYPE="HIDDEN" NAME="Tech_Links_ID" VALUE="<%=rs("tech_links_id")%>">
<INPUT TYPE="HIDDEN" NAME="Original_URL" value="<%=rs("url")%>">
<INPUT TYPE="HIDDEN" NAME="Original_Title" value="<%=rs("url_title")%>">
<INPUT TYPE="HIDDEN" NAME="Original_Desc" value="<%=rs("text")%>">
<INPUT TYPE="HIDDEN" NAME="Recipients" value="<%=Request("Recipients")%>">
<INPUT TYPE="HIDDEN" NAME="TechName" VALUE="<%=Request("TechName")%>">
<INPUT TYPE="HIDDEN" NAME="Reason" value="">

<table border=0>

<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">URL:</font></td>
	<td><INPUT TYPE="Text" NAME="URL" VALUE="<%=rs("url")%>" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Title:</font></td>
	<td><INPUT TYPE="Text" NAME="Title" VALUE="<%=rs("url_title")%>" size=40 maxlength=50></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Brief Description:</font></td>
	<td><INPUT TYPE="Text" NAME="Description" VALUE="<%=rs("text")%>" size=60 maxlength=255></td>
</tr>
</table><br>
<table align=center width=75% border=0 cellspacing=0 cellpadding=2>
	<tr>
		<td align=center>
			<input type=button class=button value="Update" OnClick='checkForm();'>
			<input type=button class=button value="Delete Link" OnClick="DeleteLink()">
			<input type=button class=button value="Cancel" OnClick='window.close();'>
			<input type=hidden NAME="Delete" value="false">
		</td>
	</tr>
</table>

</form>
</center>

<% Else %>
<script language=javascript>
<!--
var browserAbbr = "";
var browserName = navigator.appName
var browserVersion = parseInt(navigator.appVersion.substring(0,1))
if (browserName=="Microsoft Internet Explorer") {browserAbbr = "IE";}
else if (browserName=="Netscape") {browserAbbr = "NS";}

function closeWin(){i = setTimeout("window.close()", 800)}

function closeWindow(){
	if (self.opener && ((browserAbbr=="IE" && browserVersion > 3) || (browserAbbr=="NS" && browserVersion > 2))) self.opener.location.reload();
	closeWin();
	}
// -->
</script>
<% If Request("Delete") <> "true" Then %>
<html>
<head>
<title>Update Link</title>
</head>
<body bgcolor=silver onLoad="closeWindow()">
<center><font size=1 face="ms sans serif, arial, geneva">Updating information...</font></center>

<%	
	' Get submitted values
	ls_url		=	Clean(trim(Request("URL")))
	ls_title	=	Clean(trim(Request("Title")))
	ls_desc		=	Clean(trim(Request("Description")))
	tech_id		=	Request("Tech_ID")
	empname		=	Request("EmployeeName")

	set rs = DBQuery("SELECT employeeid from employee where username = '" & empname & "'")
	if not rs.eof then
		empid = rs("employeeid")
	end if
	sql =	"UPDATE tech_links " & _
			"SET url =		'" & ls_url & "', " & _
			"url_title =	'" & ls_title & "', " & _
			"text =			'" & ls_desc & "', " & _
			"username =     '" & ucase(Session("username")) & "', " & _
			"LastModified = '" & Now & "'" & _
			"WHERE tech_links_id = " & Request("Tech_Links_ID")
	set rs = DBQuery(sql)
%>

<center><font size=4 face="ms sans serif, arial, geneva">
<br><br><b>Information updated.</b></font></center>

<%Else%>

<html>
<head>
<title>Delete Link</title>
</head>
<body bgcolor=silver onLoad="closeWindow()">
<center><font size=1 face="ms sans serif, arial, geneva">Deleting Link...</font></center>

<%
' INACTIVATE TECH LINK
sql = "UPDATE tech_links SET active = 0 WHERE tech_links_id = " & Request("Tech_Links_ID")
set rs = DBQuery(sql)
' SEND EMAIL TO EXPERTS INFORMING THEM OF THE DELETION
TechName = Request("TechName")
if TechName = "" then TechName = "UNKNOWN"
BodyTxt = "I have deleted the following link:" & NL & NL & "URL: " & Request("Original_URL") & NL & "Title: " & Request("Original_Title") & NL & "Description: " & Request("Original_Desc") & NL & NL & "Reason: " & Request("Reason") & NL & NL & "Restore Link:" & NL & "http://www.sarkcolumbus.com/" & Application("Web") & "/technology/restore.asp?id=" & Request("Tech_Links_ID") & NL & NL & "(automated message sent from " & Application("Web") & ")"
SendMail "Link Deleted in '" & TechName & "'!", Request("Recipients"), "", BodyTxt
%>

<center><font size=4 face="ms sans serif, arial, geneva">
<br><br><b>Link deleted.</b></font></center>

<%
End If 'Delete 
End If %>
</body>

</html>
