<!--
Developer:	  SSeissiger
Date:         08/16/2000
Description:  Report of vacation in the next 30 days by date
			  Part of the AM report set
-->

<%
if hasRole("AccountManager") then 
%>
<!-- #include file="../section.asp" -->  
<html>
<head><title>Vacation Scheduled for the Next 30 Days</title></head>
<body>
 
<%
sql = "SELECT v.vacation_start, v.vacation_end, e.firstname, e.lastname, e.employeeID " &_
		"FROM vacation v, employee e WHERE datediff(dd, GETDATE(), vacation_start) > 0 " &_
		"AND datediff(dd, GETDATE(), vacation_start) < 30 AND v.employee_id = e.employeeid " &_
		"ORDER BY vacation_start, e.lastname, e.firstname " 
set rsVacation = DBQuery(sql)
%>
<font size=3 face='ms sans serif, arial, geneva' color=black>
<b>Vacation Scheduled for the Next 30 Days</b>
<p>
<font size=2 face="ms sans serif, arial, geneva"  color=red>
<b>Click on a hyperlink below to go directly to a Sark profile.</b>
<p>
<%if rsVacation.eof then %>
	<table><tr><font size=1 face='ms sans serif, arial, geneva'>
	<td colspan=1 width=400 align=left><font size=2 face='ms sans serif, arial, geneva' color=red>
	<b>Query Returned No Records</b>
	</td>
	</tr>
<% else %>

	<table><tr><font size=1 face='ms sans serif, arial, geneva'>
		<td colspan=1 width=140 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Employee</b></td>
		<td colspan=1 width=100 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>From</b></td>
		<td colspan=1 width=100 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>To</b></td>
	</tr>

	<%while not rsVacation.eof %>
		<tr>
		<td colspan=1 width=140 align=left><font size=1 face='ms sans serif, arial, geneva'>
			<a href='../../directory/content/details.asp?EmpID=<%=rsVacation("employeeID")%>'><%=rsVacation("LastName")%>,&nbsp;<%=rsVacation("FirstName")%></a>
		</td>
		<td colspan=1 width=100 align=left><font size=1 face='ms sans serif, arial, geneva'>
			<%=rsVacation("vacation_start")	%>
		</td>
		<td colspan=1 width=100 align=left><font size=1 face='ms sans serif, arial, geneva'>
			<%=rsVacation("vacation_end")	%>
		</td>
		</tr>
		<%rsVacation.movenext
	wend
	rsVacation.close
	%>
	
	<%end if%>
</table>
</body>
</html>

<!-- #include file="../../footer.asp" -->
	<% else %>
	<b><h2><font color=red>You are not currently an account manager,<br> therefore can not view the Account Manager Reports.
	</b></h2></font>

<%end if%>
