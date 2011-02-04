<!--
Developer:    SSMITH
Date:         12/08/1998
Description:  Edit text of a technology on tech expert page
-->

<!-- #include file="../../script.asp" -->

<html>

<head>
<title>Edit Technology Description</title>
<!-- #include file="../../style.htm" -->
</head>
<% If Request("State") = "" Then %>
<body bgcolor=silver><center>
<SCRIPT TYPE="text/javascript">
<!--
function checkForm(){
	with (document.frmInfo)
	{
		if (Description.value == "" || Description.value == null) {
			alert("You must enter a value for Description.");
			Description.focus();
			return false;
		}
	}
	document.frmInfo.submit();
	return true;
}
// -->
</SCRIPT>

<%
sql = "Select * FROM Tech WHERE Tech_ID = " & Request("Tech_ID")
Response.Write("<!-- " & sql & " -->")
set rs = DBQuery(sql)
%>
<table border=0 width='90%'><tr><td><font size=1 face="ms sans serif, arial, geneva"><b>
<%=rs("tech_name")%> Description:
</b></font></td></tr></table> 
<form NAME="frmInfo" ACTION="editdesc.asp">
<INPUT TYPE="HIDDEN" NAME="EmployeeName" VALUE="<%=Request("empname")%>">
<INPUT TYPE="HIDDEN" NAME="Tech_ID" VALUE="<%=Request("Tech_ID")%>">
<INPUT TYPE="HIDDEN" NAME="State" VALUE="Update"> 

<TEXTAREA NAME="Description" cols=60 rows=10>
<%=rs("tech_desc")%>
</TEXTAREA>
<table align=center width=75% border=0 cellspacing=0 cellpadding=2>
	<tr>
		<td align=center>
			<input type=button class=button value="Save" OnClick='checkForm();'>
			<input type=button class=button value="Cancel" OnClick='window.close();'>
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

function closeWin(){i = setTimeout("window.close()", 100)}

function closeWindow(){
	if (self.opener && ((browserAbbr=="IE" && browserVersion > 3) || (browserAbbr=="NS" && browserVersion > 2))) self.opener.location.reload();
	window.close();
	}
// -->
</script>
<html>
<head>
<title>Update Description</title>
</head>
<body bgcolor=silver onLoad="closeWindow()">
<font size=1 face="ms sans serif, arial, geneva">Updating information...</font>

<%	
	' Get submitted values
	ls_desc		=	Replace(Request("Description"),"'","''")
	tech_id		=	Request("Tech_ID")
	empname		=	Request("EmployeeName")

	set rs = DBQuery("SELECT employeeid from employee where username = '" & empname & "'")
	if not rs.eof then
		empid = rs("employeeid")
	end if
	sql =	"UPDATE tech " & _
			"SET tech_desc = '" & ls_desc & "' " & _
			"WHERE tech_id = " & Request("Tech_ID")
	'Debug Output
	'Response.Write(sql)

	set rs = DBQuery(sql)
End If 
%>
</body>

</html>
