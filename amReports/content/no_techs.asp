<% @language=VBscript%>

<!--
Developer:	  DMartin
Date:         10/09/2000
Description:  Report of  no technologies listed
-->
<%'if (hasRole("AccountManager")) then %>
<!-- #include file="../section.asp" -->

<html>
<head><title>Skills Inventory -- Not Entered on Intranet</title></head>

<body>

<font size=3 face='ms sans serif, arial, geneva' color=blue>
<b>Technologies -- None Entered on Intranet</b>
<br></br>
<%

	  
sql = "SELECT firstname, lastname, EmployeeID, Employee_ID, EmployeeTitle_ID " &_ 
	  "FROM Employee LEFT OUTER JOIN Employee_Tech_Xref " &_
	  "ON Employee_ID = EmployeeID " &_
      "WHERE employee_ID IS NULL " &_
      "AND EmployeeTitle_ID in (1, 2, 3, 7, 8, 23) " &_
      "ORDER BY lastname, firstname"
	  
	  
	  
	  
	  
	  
	  'Response.Write"(" & sql & ")"
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
	<td colspan=1 width=140 align=left><font size=2 face="ms sans serif, arial, geneva" color=blue><b>Employee</b></td>
</tr>

<%while not rsEmployee.eof 

%>
	<tr>
	<td colspan=1 width=140 align=left><font size=1 face='ms sans serif, arial, geneva'>
		<a href='../../directory/content/details.asp?EmpID=<%=rsEmployee("employee_ID")%>'><%=rsEmployee("LastName")%>,&nbsp;<%=rsEmployee("FirstName")%></a>
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
<% 'else %>
	<!--<b><h2><font color=red>You are not currently an account manager,<br> therefore can not view the Account Manager Reports.
	</b></h2></font>
-->
<%'end if%>