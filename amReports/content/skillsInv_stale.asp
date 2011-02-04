<% @language=VBscript%>

<!--
Developer:	  SSeissiger
Date:         08/14/2000
Description:  Report of skills inv
Changed By:	  Julie Walters
Date Changed: 11/06/2000
-->

<%if (hasRole("AccountManager")) then %>
<!-- #include file="../section.asp" -->
<html>
<head><title>Skills Inventory Last Modified by Sark</title></head>

<body>
<font size=2 face="ms sans serif, arial, geneva"  color=red>
<b>Click on a hyperlink below to go directly to a Sark profile.</b>
<font size=3 face='ms sans serif, arial, geneva' color=black>
<p>
<b>Skills Inventory -- 90 to 104 Days Old</b>
<p>
<%
sql = "SELECT FirstName, LastName, employeeID, DateSkillsModified FROM Employee " &_
		"WHERE DATEDIFF(dd, DateSkillsModified, GETDATE()) > 89 " &_
		"AND DATEDIFF(dd, DateSkillsModified, GETDATE()) < 105 " &_
		"AND employeetitle_ID in (1,2,3,4,5,7,8,19,23,20) " &_
		"ORDER BY LastName"
	
set rsEmployee = DBQuery(sql)
%>

<%if rsEmployee.eof then %>
	<table><tr><font size=1 face='ms sans serif, arial, geneva'>
	<td colspan=1 width=400 align=left><font size=2 face='ms sans serif, arial, geneva' color=blue>
	<b>Query Returned No Records</b>
	</td>
	</tr>
<% else %>

<table><tr><font size=1 face='ms sans serif, arial, geneva'>
	<td colspan=1 width=140 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Employee</b></td>
	<td colspan=1 width=100 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Last Modified</b></td>
</tr>

<%while not rsEmployee.eof %>
	<tr>
	<td colspan=1 width=140 align=left><font size=1 face='ms sans serif, arial, geneva'>
		<a href='../../directory/content/details.asp?EmpID=<%=rsEmployee("employeeID")%>'><%=rsEmployee("LastName")%>,&nbsp;<%=rsEmployee("FirstName")%></a>
	</td>
	<td colspan=1 width=100 align=left><font size=1 face='ms sans serif, arial, geneva'>
		<% =rsEmployee("DateSkillsModified")%>
	</td>
	</tr>
		
<%	rsEmployee.movenext
	wend
	rsEmployee.close
%>
	<%end if%>
	
</table>

<p>
<font size=3 face='ms sans serif, arial, geneva' color=black>
<b>Skills Inventory -- 105 to 119 Days Old</b>
<p>
<%
sql = "SELECT FirstName, LastName, employeeID, DateSkillsModified FROM Employee " &_
		"WHERE DATEDIFF(dd, DateSkillsModified, GETDATE()) > 104 " &_
		"AND DATEDIFF(dd, DateSkillsModified, GETDATE()) < 120 " &_
		"AND employeetitle_ID in (1,2,3,4,5,7,8,19,23,20) " &_
		"ORDER BY LastName"
	
set rsEmployee = DBQuery(sql)
%>

<%if rsEmployee.eof then %>
	<table><tr><font size=1 face='ms sans serif, arial, geneva'>
	<td colspan=1 width=400 align=left><font size=2 face='ms sans serif, arial, geneva' color=blue>
	<b>Query Returned No Records</b>
	</td>
	</tr>
<% else %>

<table><tr><font size=1 face='ms sans serif, arial, geneva'>
	<td colspan=1 width=140 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Employee</b></td>
	<td colspan=1 width=100 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Last Modified</b></td>
</tr>

<%while not rsEmployee.eof %>
	<tr>
	<td colspan=1 width=140 align=left><font size=1 face='ms sans serif, arial, geneva'>
		<a href='../../directory/content/details.asp?EmpID=<%=rsEmployee("employeeID")%>'><%=rsEmployee("LastName")%>,&nbsp;<%=rsEmployee("FirstName")%></a>
	</td>
	<td colspan=1 width=100 align=left><font size=1 face='ms sans serif, arial, geneva'>
		<% =rsEmployee("DateSkillsModified")%>
	</td>
	</tr>
		
<%	rsEmployee.movenext
	wend
	rsEmployee.close
%>
	<%end if%>
	
</table>

<p>
<font size=3 face='ms sans serif, arial, geneva' color=black>
<b>Skills Inventory -- 120+ Days Old</b>
<p>
<%
sql = "SELECT FirstName, LastName, employeeID, DateSkillsModified FROM Employee " &_
		"WHERE DATEDIFF(dd, DateSkillsModified, GETDATE()) > 119 " &_
		"AND employeetitle_ID in (1,2,3,4,5,7,8,19,23,20) " &_
		"ORDER BY LastName"
	
set rsEmployee = DBQuery(sql)
%>

<%if rsEmployee.eof then %>
	<table><tr><font size=1 face='ms sans serif, arial, geneva'>
	<td colspan=1 width=400 align=left><font size=2 face='ms sans serif, arial, geneva' color=blue>
	<b>Query Returned No Records</b>
	</td>
	</tr>
<% else %>

<table><tr><font size=1 face='ms sans serif, arial, geneva'>
	<td colspan=1 width=140 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Employee</b></td>
	<td colspan=1 width=100 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Last Modified</b></td>
</tr>

<%while not rsEmployee.eof %>
	<tr>
	<td colspan=1 width=140 align=left><font size=1 face='ms sans serif, arial, geneva'>
		<a href='../../directory/content/details.asp?EmpID=<%=rsEmployee("employeeID")%>'><%=rsEmployee("LastName")%>,&nbsp;<%=rsEmployee("FirstName")%></a>
	</td>
	<td colspan=1 width=100 align=left><font size=1 face='ms sans serif, arial, geneva'>
		<% =rsEmployee("DateSkillsModified")%>
	</td>
	</tr>
		
<%	rsEmployee.movenext
	wend
	rsEmployee.close
%>
	<%end if%>
	
</table>

</body>
</html>

<!-- #include file="../../footer.asp" -->
<% else %>
	<b><h2><font color=red>You are not currently an account manager, therefore can not view the Account Manager Reports.
	</b></h2></font>

<%end if%>