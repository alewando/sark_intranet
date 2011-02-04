<!--
Developer:	  JWalters
Date:         10/24/2000
Description:  Report of exams passed in 0-30 days and 31-60 days back.
			  Part of the AM report set
-->

<%
if hasRole("AccountManager") or hasRole("Sales") then 
%>
<!-- #include file="../section.asp" -->  
<html>
<head><title>Exams Taken in the Past 60 days</title></head>
<body>
<p>
<font size=2 face="ms sans serif, arial, geneva"  color=red>
<b>Click on a hyperlink below to go directly to a Sark profile.</b>
<p>
<font size=3 face='ms sans serif, arial, geneva' color=black>
<b>Exams Taken in the Past 30 Days</b>
<p>
<%
sql = "SELECT e.firstname, e.lastname, ep.exam_date, el.exam_name, e.employeeID " &_
		"FROM exams_passed ep, exam_list el, employee e " &_
		"WHERE DATEDIFF(dd, GETDATE(), exam_date) > -31 " &_
		"AND DATEDIFF(dd, GETDATE(), exam_date) < 0 " &_
		"AND e.employeeid = ep.employeeid AND ep.exam_id = el.exam_id " &_
		"ORDER BY e.lastname, e.firstname, ep.exam_date "

set rsExamPassed = DBQuery(sql)
%>
<% if rsExamPassed.eof then %>
	<table><tr><font size=1 face='ms sans serif, arial, geneva'>
	<td colspan=1 width=400 align=left><font size=2 face='ms sans serif, arial, geneva' color=red>
	<b>Query Returned No Records</b>
	</td>
	</tr>
<% else %>

	<table><tr><font size=1 face='ms sans serif, arial, geneva'>
		<td colspan=1 width=140 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Employee</b></td>
		<td colspan=1 width=100 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Pass Date</b></td>
		<td colspan=1 width=100 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Exam Name</b></td>
	</tr>

	<%while not rsExamPassed.eof %>
		<tr>
		<td colspan=1 width=140 align=left><font size=1 face='ms sans serif, arial, geneva'>
			<a href='../../directory/content/details.asp?EmpID=<%=rsExamPassed("employeeID")%>'><%=rsExamPassed("LastName")%>,&nbsp;<%=rsExamPassed("FirstName")%></a>
		</td>
		<td colspan=1 width=100 align=left><font size=1 face='ms sans serif, arial, geneva'>
			<%=rsExamPassed("exam_date")	%>
		</td>
		<td colspan=1 width=250 align=left><font size=1 face='ms sans serif, arial, geneva'>
			<%=rsExamPassed("exam_name")	%>
		</td>
		</tr>
		<%rsExamPassed.movenext
	wend
	rsExamPassed.close
	%>
	
	<%end if%>
</table>

<p>
<font size=3 face='ms sans serif, arial, geneva' color=black>
<b>Exams Taken in the Past 31-60 Days</b>
<p>
<%
sql = "SELECT e.firstname, e.lastname, ep.exam_date, el.exam_name, e.employeeID " &_
		"FROM exams_passed ep, exam_list el, employee e " &_
		"WHERE DATEDIFF(dd, GETDATE(), exam_date) > -61 " &_
		"AND DATEDIFF(dd, GETDATE(), exam_date) < -30 " &_
		"AND e.employeeid = ep.employeeid AND ep.exam_id = el.exam_id " &_
		"ORDER BY e.lastname, e.firstname, ep.exam_date "

set rsExamPassed = DBQuery(sql)
%>
<% if rsExamPassed.eof then %>
	<table><tr><font size=1 face='ms sans serif, arial, geneva'>
	<td colspan=1 width=400 align=left><font size=2 face='ms sans serif, arial, geneva' color=red>
	<b>Query Returned No Records</b>
	</td>
	</tr>
<% else %>

	<table><tr><font size=1 face='ms sans serif, arial, geneva'>
		<td colspan=1 width=140 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Employee</b></td>
		<td colspan=1 width=100 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Pass Date</b></td>
		<td colspan=1 width=100 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Exam Name</b></td>
	</tr>

	<%while not rsExamPassed.eof %>
		<tr>
		<td colspan=1 width=140 align=left><font size=1 face='ms sans serif, arial, geneva'>
			<a href='../../directory/content/details.asp?EmpID=<%=rsExamPassed("employeeID")%>'><%=rsExamPassed("LastName")%>,&nbsp;<%=rsExamPassed("FirstName")%></a>
		</td>
		<td colspan=1 width=100 align=left><font size=1 face='ms sans serif, arial, geneva'>
			<%=rsExamPassed("exam_date")	%>
		</td>
		<td colspan=1 width=250 align=left><font size=1 face='ms sans serif, arial, geneva'>
			<%=rsExamPassed("exam_name")	%>
		</td>
		</tr>
		<%rsExamPassed.movenext
	wend
	rsExamPassed.close
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
