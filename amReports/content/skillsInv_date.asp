<% @language=VBscript%>

<!--
Developer:	  SSeissiger
Date:         08/14/2000
Description:  Report of skills inv
-->
<%if (hasRole("AccountManager")) then %>
<!-- #include file="../section.asp" -->

<html>
<head><title>Skills Inventory Last Modified by Sark</title></head>

<body>
<font size=3 face='ms sans serif, arial, geneva' color=black>
<b>Skills Inventory -- Date Last Modified</b>
<p>
<font size=2 face="ms sans serif, arial, geneva"  color=red>
<b>Click on a hyperlink below to go directly to a Sark profile.</b>
<p>
<%
sql = "SELECT FirstName, LastName, EmployeeID, DateSkillsModified FROM Employee WHERE employeetitle_ID in (1,2,3,4,5,7,8,19,23,20)ORDER BY LastName"
set rsEmployee = DBQuery(sql)
%>
<table><tr><font size=1 face='ms sans serif, arial, geneva'>
	<td colspan=1 width=140 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Employee</b></td>
	<td colspan=1 width=100 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Last Modified</b></td>
</tr>

<%while not rsEmployee.eof 
	dateModified = rsEmployee("DateSkillsModified")
%>
	
	<tr>
	<td colspan=1 width=140 align=left><font size=1 face='ms sans serif, arial, geneva'>
		<a href='../../directory/content/details.asp?EmpID=<%=rsEmployee("employeeID")%>'><%=rsEmployee("LastName")%>,&nbsp;<%=rsEmployee("FirstName")%></a>
	</td>
	<td colspan=1 width=100 align=left><font size=1 face='ms sans serif, arial, geneva'>
		<% if not isnull(dateModified)then
			Response.write(dateModified)
			else Response.Write("<b>None Entered</b>")
			end if
		%>
	</td>
	</tr>
<%	rsEmployee.movenext
	wend
	rsEmployee.close
%>
</table>
</body>
</html>

<!-- #include file="../../footer.asp" -->
<% else %>
	<b><h2><font color=red>You are not currently an account manager therefore can not view the Account Manager Reports.
	</b></h2></font>

<%end if%>