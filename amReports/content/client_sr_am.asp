<!--
Developer:	  SSeissiger
Date:         08/17/2000
Description:  Report - client list with sales rep and account manager
			  Part of the AM report set
-->
<% @language=VBscript%>
<%if (hasRole("AccountManager")) then %>
<!-- #include file="../section.asp" -->

<html>
<head><title>Client List with Sales Rep and Account Manager</title></head>

<body>
<%
sql = "select c.clientname,sr.firstname as srFirstName, sr.lastname as srLastName, sr.employeeID as srEmpID, " &_
"am.firstname as amFirstName, am.lastname as amLastName, am.EmployeeID as amEmpID " &_
"from client c left join employee sr on " &_
"c.Employee_SalesRep_ID = sr.employeeID " &_
"left join employee am on " &_
"c.Employee_AcctMgr_ID = am.employeeID " &_
"where clientname not in ('Unassigned', 'Office Staff') " &_
"ORDER BY clientname"

set rs = DBQuery(sql)
%>
<font size=3 face='ms sans serif, arial, geneva' color=black>
<b>Client List with Sales Rep and Acct Mgr</b>
<p>
<font size=2 face="ms sans serif, arial, geneva"  color=red>
<b>Click on a hyperlink below to go directly to a Sark profile.</b>
<p>
	<table><tr><font size=1 face='ms sans serif, arial, geneva'>
		<td colspan=1 width=210 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Client</b></td>
		<td colspan=1 width=125 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Sales Rep</b></td>
		<td colspan=1 width=125 align=left><font size=2 face="ms sans serif, arial, geneva" color=maroon><b>Acct Mgr</b></td>
	</tr>

	<%while not rs.eof
		srLastName = rs("srLastName")
		srFirstName = rs("srFirstName")
		srEmpID = rs("srEmpID")
		amLastName = rs("amLastName")
		amFirstName = rs("amFirstName")
		amEmpID = rs("amEmpID")
	%>
		<tr>
		<td colspan=1 width=210 align=left><font size=1 face='ms sans serif, arial, geneva'>
			<%=rs("clientname")	%>
		</td>
		<td colspan=1 width=125 align=left><font size=1 face='ms sans serif, arial, geneva'>
			<%if not isnull(srLastName) then %>
				<a href='../../directory/content/details.asp?EmpID=<%=srEmpID%>'><%=srLastName%>,&nbsp;<%=srFirstName%></a><br>				</td>				<%else%>&nbsp;
				<%end if%>
		</td>
		
		<td colspan=1 width=125 align=left><font size=1 face='ms sans serif, arial, geneva'>
			<%if not isnull(amLastName) then %>
				<a href='../../directory/content/details.asp?EmpID=<%=amEmpID%>'><%=amLastName%>,&nbsp;<%=amFirstName%></a><br>				</td>				<%else%>&nbsp;
				<%end if%>
		</td>
		
		</tr>
		<%rs.movenext
	wend
	rs.close
	%>
</table>
</body>
</html>

<!-- #include file="../../footer.asp" -->
<% else %>
	<b><h2><font color=red>You are not currently an account manager,<br> therefore can not view the Account Manager Reports.
	</b></h2></font>

<%end if%>
