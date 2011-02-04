<!--
Developer:    SSMITH
Date:         11/30/1998
Description:  Add link to tech expert page
-->

<!-- #include file="../../script.asp" -->

<html>

<head>
<title>Add Technology Link</title>
<!-- #include file="../../style.htm" -->
</head>


<% If Request("State") = "" Then %>
<body bgcolor=silver onLoad="document.frmInfo.URL.focus()"><center>

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
	if (confirm("The system will automatically notify the 'experts' and alert them to the adding of this link.")){
		document.frmInfo.submit();
		return true;
		}
	return false;
	}
// -->
</script>

<table border=0 width='90%'><tr><td><font size=1 face="ms sans serif, arial, geneva"><b>
Enter the link information:
</b></font></td></tr></table> 


<form name="frmInfo" method="post" ACTION="links_add.asp">
<INPUT TYPE="HIDDEN" NAME="EmployeeName" VALUE="<%=Request("empname")%>">
<INPUT TYPE="HIDDEN" NAME="Tech_ID" VALUE="<%=Request("Tech_ID")%>">
<INPUT TYPE="HIDDEN" NAME="State" VALUE="Update">
<INPUT TYPE="HIDDEN" NAME="Recipients" VALUE="<%=Request("Recipients")%>">
<INPUT TYPE="HIDDEN" NAME="TechName" VALUE="<%=Request("TechName")%>">

<table border=0>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">URL:</font></td>
	<td><INPUT TYPE="Text" NAME="URL" VALUE="http://" size=40 maxlength=255></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Title:</font></td>
	<td><INPUT TYPE="Text" NAME="Title" VALUE="" size=40 maxlength=50></td>
</tr>
<tr>
	<td align=right><font size=1 face="ms sans serif, arial, geneva">Brief Description:</font></td>
	<td><INPUT TYPE="Text" NAME="Description" VALUE="" size=60 maxlength=255></td>
</tr>
</table><br>
<table align=center width="75%" border=0 cellspacing=0 cellpadding=2>
	<tr>
		<td align=center>
			<input type=button class=button value="Save" onClick="checkForm()">
			<input type=button class=button value="Cancel" onClick="window.close()">
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

function closeWindow(){
	if (self.opener && ((browserAbbr=="IE" && browserVersion > 3) || (browserAbbr=="NS" && browserVersion > 2))) self.opener.location.reload();
	window.close();
	}
// -->
</script>


<html>
<head>
<title>Update Link Information</title>
</head>
<body bgcolor=silver onLoad="closeWindow()">
<center><font size=1 face="ms sans serif, arial, geneva">Saving information...</font></center>

<%	
	' Get submitted values
	ls_url		=	Clean(trim(Request("URL")))
	ls_title	=	Clean(trim(Request("Title")))
	ls_desc		=	Clean(trim(Request("Description")))
	tech_id		=	Request("Tech_ID")
	empname		=	ucase(Request("EmployeeName"))
	empid		=	""

	set rs = DBQuery("SELECT employeeid from employee where username = '" & empname & "'")
	if not rs.eof then
		empid = rs("employeeid")
		sql = "Insert into tech_links (tech_id, url, url_title, text, employee_id, active, username, num_hits,timestamp) values (" & tech_id & ", '" & ls_url & "', '" & ls_title & "', '" & ls_desc & "', " & empid &", 1, '" & empname & "', 0, '" & Now & "')"
	else
		sql = "Insert into tech_links (tech_id, url, url_title, text, active, username, num_hits,timestamp) values (" & tech_id & ", '" & ls_url & "', '" & ls_title & "', '" & ls_desc & "', 1, '" & empname & "', 0, '" & Now & "')"
	end if
	set rs = DBQuery(sql)

	' SEND EMAIL TO EXPERTS INFORMING THEM OF THE DELETION
	TechName = Request("TechName")
	if TechName = "" then TechName = "UNKNOWN"
	BodyTxt = "I have added the following link"
	if empid = "" then BodyTxt = BodyTxt & " from a guest account"
	BodyTxt = BodyTxt & ":" & NL & NL & "URL: " & Request("URL") & NL & "Title: " & Request("Title") & NL & "Description: " & Request("Description") & NL & NL & "(automated message sent from " & Application("Web") & ")"
	SendMail "Link Added to '" & TechName & "'!", Request("Recipients"), "", BodyTxt
End If %>
</body>

</html>
