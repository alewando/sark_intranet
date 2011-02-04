<% @language=VBscript%>

<!--
Developer:	  SSeissiger
Date:         08/14/2000
Description:  Report of skills inv
-->
<%if (hasRole("AccountManager")) then %>
<!-- #include file="../section.asp" -->

<html>
<head><title>Skills Inventory -- Not Entered on Intranet</title></head>

<body>

<font size=3 face='ms sans serif, arial, geneva' color=black>
<b>Skills Inventory -- Not Entered on Intranet</b>
<p>
<font size=2 face="ms sans serif, arial, geneva"  color=red>
<b>Click on a hyperlink below to go directly to a Sark profile.</b>
<p>
<%
sql = "SELECT FirstName, LastName, employeeID, DateSkillsModified FROM Employee WHERE employeetitle_ID in (1,2,3,4,5,7,8,19,23,20) " &_
	"ORDER BY LastName"
set rsEmployee = DBQuery(sql)
%>

<%if rsEmployee.eof then %>
	<table><tr><font size=1 face='ms sans serif, arial, geneva'>
	<td colspan=1 width=400 align=left><font size=2 face='ms sans serif, arial, geneva' color=red>
	<b>Query Returned No Records</b>
	</td>
	</tr>
<% else %>

<table><tr><font size=1 face='ms sans serif, arial, geneva'>
	<td colspan=1 width=140 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Employee</b></td>
</tr>

<%while not rsEmployee.eof 
	dateModified = rsEmployee("DateSkillsModified")
%>
	<% if isnull(dateModified)then %>
	<tr>
	<td colspan=1 width=140 align=left><font size=1 face='ms sans serif, arial, geneva'>
		<a href='../../directory/content/details.asp?EmpID=<%=rsEmployee("employeeID")%>'><%=rsEmployee("LastName")%>,&nbsp;<%=rsEmployee("FirstName")%></a>
	</td>
	</tr>
		
	<% end if%>
<%	rsEmployee.movenext
	wend
	rsEmployee.close
%>

<% end if%>
</table>
</body>
</html>

<!-- #include file="../../footer.asp" -->
<% else %>
	<b><h2><font color=red>You are not currently an account manager therefore can not view the Account Manager Reports.
	</b></h2></font>

<%end if%>