<% response.expires = 0 %>
<!-- #include file="../../script.asp" -->


<%
'---------------------------------------------------------------------------------
' Description:	Adds sarks to a project.
' History:		06/10/1999 - KDILL - Created
'---------------------------------------------------------------------------------
%>


<html>
<head>
<title>Add/Remove "<% =request("ProjectName") %>" Sarks</title>
<!-- #include file="../../style.htm" -->
</head>


<body bgcolor=silver>
<font face="ms sans serif, arial, geneva">
<center>

<%
ProjectID = CInt(request("ProjectID"))
if request("action") = "update" then
%>

<font size=3 face="ms sans serif, arial, geneva">
Saving...
</font>

<%
	'-----------------------------------------------
	'	Update project employee listing				
	'-----------------------------------------------
	DBQuery("DELETE FROM Employee_Project_Xref WHERE Project_ID = " & ProjectID)
	for each x in request("Sarks")
		DBQuery("INSERT INTO Employee_Project_Xref (Project_ID, Employee_ID) values (" & ProjectID & ", " & x & ")")
		next
%>

<script language=javascript>
<!--
var browserAbbr = "";
var browserName = navigator.appName
var browserVersion = parseInt(navigator.appVersion.substring(0,1))
if (browserName=="Microsoft Internet Explorer") {browserAbbr = "IE";}
else if (browserName=="Netscape") {browserAbbr = "NS";}
function closeWin(){
	if (self.opener && ((browserAbbr=="IE" && browserVersion > 3) || (browserAbbr=="NS" && browserVersion > 2))) self.opener.location.reload();
	window.close()
	}
i = setTimeout("closeWin()", 700)
// -->
</script>

<% else %>

<table border=0 width='90%'><tr><td><font size=1 face="ms sans serif, arial, geneva">
<b>Please select the Sarks who worked on the project.</b>
</font></td></tr></table>

<form name="frmInfo" method=post action="details_add_emps.asp">
<input type=hidden name="ProjectID" value="<%=request("ProjectID") %>">
<input type=hidden name="ProjectName" value="<% =request("ProjectName") %>">
<input type=hidden name="ClientID" value="<% =request("ClientID") %>">

<table border=0><tr><td><font size=1 face='ms sans serif, arial, geneva'>
<%
'-------------------------------------------------
'	Execute Database Query for Tech Listing Info  
'-------------------------------------------------
sql = "SELECT Employee.EmployeeID, Employee.FirstName, Employee.LastName, Employee_Project_Xref.Project_ID FROM Employee LEFT OUTER JOIN Employee_Project_Xref ON Employee.EmployeeID = Employee_Project_Xref.Employee_ID AND Employee.DateTermination IS NULL AND Employee.Branch_ID = " & Application("DefaultBranchID") & " ORDER BY Employee.LastName, Employee.FirstName"
set rs = DBQuery(sql)
i = 0
while not rs.eof
	if (i <> 0) and (i <> rs("EmployeeID")) then
		Response.Write("<INPUT TYPE=checkbox NAME=sarks VALUE=" & sark_id & " " & checked & ">" & sark_name & "<BR>" & NL)
		checked = ""
		end if
	sark_id = rs("EmployeeID")
	sark_name = rs("LastName") & ", " & rs("FirstName")
	i = sark_id
	if not isnull(rs("Project_ID")) then
		if CInt(rs("Project_ID")) = ProjectID then
			checked = "checked"
			end if
		end if
	rs.movenext
	wend
rs.close
Response.Write("<INPUT TYPE=checkbox NAME=sarks VALUE=" & sark_id & " " & checked & ">" & sark_name & "<BR>" & NL)
set rs = nothing
%>
</td></tr></table>

<input type=hidden name="action" value="update"><br>
<table align=center width=75% border=0 cellspacing=0 cellpadding=2>
	<tr>
		<td align=center>
			<input type=button class=button value="Save Sarks" OnClick='document.frmInfo.submit();'>
			<input type=button class=button value="Cancel" OnClick='window.close();'>
		</td>
	</tr>
</table>
</form>

<% end if %>

</center>
</body>
</html>

<%
on error resume next
DataConn.close
set DataConn = nothing
%>
