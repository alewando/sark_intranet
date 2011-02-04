<!-- #include file="../section.asp" -->


<%
'---------------------------------------------------------------------------------
' Description:	Statistics about the office.
' History:		10/29/1999 - KDILL - Created
'               11/18/1999 - DKATHAN - Removed Oldest, Youngest, Avg Ages
'				01/13/2000 - KDILL - Added Technology stats
'---------------------------------------------------------------------------------
%>



<center>
<table border=0 cellpadding=10 bgcolor=#ffffcc class=tableShadow><tr><td>

<table width="100%" border=0 cellpadding=3>


	<tr><td colspan=3><font size=1 face="ms sans serif, arial, geneva">
		<b>Branch Statistics</b>
		</font></td>



<% set rs = DBQuery("SELECT count(EmployeeID) FROM employee") 
   totalEmps = rs(0)
%>

	<tr><td width=20></td><td><font size=1 face="ms sans serif, arial, geneva">Total Employees</font></td>	
		<td><font size=1 face="ms sans serif, arial, geneva"><% =rs(0) %></font></td>
		</tr>
<% rs.close %>

<% ' removed Total Consultants because total is incorrect by using VM extensions %>
<% ' removed Avg Sark Age and Avg Sark Years %>

	<tr><td colspan=3><font size=1 face="ms sans serif, arial, geneva">
		<br><b>Sark Statistics</b>
		</font></td>



<% set rs = DBQuery("SELECT EmployeeID, FirstName, LastName, StartDate FROM employee ORDER BY StartDate") %>
	<tr><td width=10></td><td valign=top><font size=1 face="ms sans serif, arial, geneva">Oldtimer SARKs</font></td>	
		<td><font size=1 face="ms sans serif, arial, geneva">
<% for i = 1 to 3 %>
			<a href="../../directory/content/details.asp?EmpID=<% =rs("EmployeeID") %>"><% =rs("FirstName") & " " & rs("LastName") %></a>&nbsp;&nbsp;&nbsp;<font color=gray>(<% =DateDiff("yyyy", rs("StartDate"), Date) %> years)</font><br>
<%
	rs.movenext
next
%>
			</font></td>	
		</tr>
<% rs.close %>

<% set rs = DBQuery("SELECT EmployeeID, FirstName, LastName, StartDate FROM employee ORDER BY StartDate DESC") %>
	<tr><td width=10></td><td valign=top><font size=1 face="ms sans serif, arial, geneva">Newest SARKs</font></td>	
		<td><font size=1 face="ms sans serif, arial, geneva">
<% for i = 1 to 3 %>
			<a href="../../directory/content/details.asp?EmpID=<% =rs("EmployeeID") %>"><% =rs("FirstName") & " " & rs("LastName") %></a>&nbsp;&nbsp;&nbsp;<font color=gray>(<% =rs("StartDate") %>)<br>
<%
	rs.movenext
next
%>
			</font></td>
		</tr>
<% rs.close %>

<% ' Average Age 
set rs = DBQuery("SELECT SUM(DATEDIFF(yy, Birthday, GETDATE())) AS Age FROM Employee WHERE DATEDIFF(yy, Birthday, GETDATE()) < 90")
totalAge = rs(0)
set rs = DBQuery("SELECT COUNT(*) FROM Employee WHERE DATEDIFF(yy, Birthday, GETDATE()) < 90")
empCount=rs(0)
avgAge = FormatNumber(totalAge/empCount,1)

%>
	<tr><td width=10></td><td valign=top><font size=1 face="ms sans serif, arial, geneva">Average Age</font></td>	
		<td><font size=1 face="ms sans serif, arial, geneva">
          <%=avgAge%>
        </font></td>
    </tr>

<% set rs = DBQuery("SELECT City, COUNT(City) AS cnt FROM Employee WHERE Branch_ID = " & Application("DefaultBranchID") & " AND DateTermination is null GROUP BY City ORDER BY cnt DESC") %>
	<tr><td width=10></td><td valign=top><font size=1 face="ms sans serif, arial, geneva">Most popular place to live</font></td>	
		<td><font size=1 face="ms sans serif, arial, geneva">
<% for i = 1 to 4 %>
			<% =i %>.&nbsp;&nbsp;&nbsp;<% =rs("City") %>&nbsp;&nbsp;&nbsp;<font color=gray>(<% =rs("Cnt") %>)</font><br>
<%
	rs.movenext
next
%>
			</font></td>
		</tr>
<% rs.close %>



	<tr><td colspan=3><font size=1 face="ms sans serif, arial, geneva">
		<br><b>Client Statistics</b>
		</font></td>



<% set rs = DBQuery("SELECT COUNT(DISTINCT Client.Client_ID) FROM Client INNER JOIN Employee ON Client.Client_ID = Employee.Client_ID WHERE Client.SortOverride = 0 AND Client.Client_ID > 2") %>
	<tr><td width=10></td><td><font size=1 face="ms sans serif, arial, geneva">Total Active Clients</font></td>	
		<td><font size=1 face="ms sans serif, arial, geneva"><% =rs(0) %></font></td>
		</tr>
<% rs.close %>



<% set rs = DBQuery("SELECT Client.ClientName, COUNT(Employee.EmployeeID) AS cnt FROM Client INNER JOIN Employee ON Client.Client_ID = Employee.Client_ID WHERE (Employee.Branch_ID = " & Application("DefaultBranchID") & ") AND (Employee.DateTermination IS NULL) AND Client.SortOverride = 0 AND Client.Client_ID > 2 GROUP BY Client.ClientName ORDER BY cnt DESC") %>
	<tr><td width=10></td><td valign=top><font size=1 face="ms sans serif, arial, geneva">Top Active Clients</font></td>	
		<td><font size=1 face="ms sans serif, arial, geneva">
<% for i = 1 to 4 %>
			<% =i %>.&nbsp;&nbsp;&nbsp;<% =rs("ClientName") %>&nbsp;&nbsp;&nbsp;<font color=gray>(<% =rs("Cnt") %>)</font><br>
<%
	rs.movenext
next
%>
			</font></td>
		</tr>
<% rs.close %>

	<tr><td colspan=3><font size=1 face="ms sans serif, arial, geneva">
		<br><b>Exams / Certifications</b>
		</font></td>
	<% set rs = DBQuery("SELECT employeeid, firstname, lastname, " &_
	             "COUNT(Employee_Cert_Xref) AS ct " & _
                 "FROM Employee e, Employee_Cert_Xref ecx " & _
                 "WHERE e.employeeid = ecx.employee_ID " & _
                 "GROUP BY employeeid, lastname, firstname " & _
                 "ORDER BY ct DESC") 
    %>
	<tr><td width=10></td><td valign=top><font size=1 face="ms sans serif, arial, geneva">Most Certifications</font></td>	
		<td><font size=1 face="ms sans serif, arial, geneva">
		<%
		certCount=1
		if not rs.eof then
		 lastCount=rs("ct")
		end if
		do while not rs.eof and certCount<=1
		 if lastCount<>rs("ct") then
			certCount = certCount+1
		 end if
		 Response.Write(certCount & ".&nbsp;&nbsp;&nbsp;")
		 Response.Write("<A HREF='../../directory/content/details.asp?EmpID=" & rs("employeeid") & "'>")
		 Response.Write(rs("firstName") & " " & rs("lastName"))
		 Response.Write("</A>")
		 Response.Write("&nbsp;&nbsp;&nbsp;")
		 Response.Write("<font color=gray>(" & rs("ct") & ")</font><br>")
		 lastCount = rs("ct")
		 rs.moveNext
		loop
		%>
        </font></td>
    </tr>
    
	<% set rs = DBQuery("SELECT e.employeeid, firstname, lastname, " & _
                 "COUNT(Exams_passed_ID) AS ct " & _
                 "FROM Employee e, Exams_passed ep " & _
                 "WHERE e.employeeid = ep.employeeID " & _
                 "GROUP BY e.employeeid, lastname, firstname " & _
                 "ORDER BY ct DESC") 
    %>
	<tr><td width=10></td><td valign=top><font size=1 face="ms sans serif, arial, geneva">
	 Most Exams</font></td>	
		<td><font size=1 face="ms sans serif, arial, geneva">
		<%
		examCount=1
		if not rs.eof then
		 lastCount=rs("ct")
		end if
		do while not rs.eof and rs("ct")=lastCount
		 Response.Write(examCount & ".&nbsp;&nbsp;&nbsp;")
		 Response.Write("<A HREF='../../directory/content/details.asp?EmpID=" & rs("employeeid") & "'>")
		 Response.Write(rs("firstName") & " " & rs("lastName"))
		 Response.Write("</A>")
		 Response.Write("&nbsp;&nbsp;&nbsp;")
		 Response.Write("<font color=gray>(" & rs("ct") & ")</font>")
		 lastCount = rs("ct")
		 rs.moveNext
		loop
		%>
        </font></td>
    </tr>    

	<% set rs=DBQuery("SELECT Employee.EmployeeID, " & _
           "FirstName + ' ' + LastName AS EmpName, " &_
           "Exam_Name, Exam_date " & _
           "FROM Exams_passed INNER JOIN " &_
           "Employee ON " &_
           "Exams_passed.EmployeeID = Employee.EmployeeID INNER JOIN " & _
           "Exam_List ON " & _
           "Exams_passed.Exam_ID = Exam_List.Exam_ID " & _
           "ORDER BY Exam_Date DESC") 
	%>
	<tr><td width=10></td><td valign=top><font size=1 face="ms sans serif, arial, geneva">Most Recent Exam</font></td>
		<td><font size=1 face="ms sans serif, arial, geneva">
	    <A HREF="../../directory/content/details.asp?EmpID=<%=rs("EmployeeID")%>"><%=rs("EmpName")%></A><BR>
	    (<%=rs("Exam_Name")%> on <%=rs("Exam_date")%>)<BR>	
		</font>
		</td>
	</tr>


	<tr><td colspan=3><font size=1 face="ms sans serif, arial, geneva">
		<br><b>Recruiting Statistics</b>
		</font></td>



<% set rs = DBQuery("SELECT College, COUNT(College) AS cnt FROM Profile WHERE College <> '' AND College IS NOT NULL GROUP BY College ORDER BY cnt DESC") %>
	<tr><td width=10></td><td valign=top><font size=1 face="ms sans serif, arial, geneva">Top Recruiting Schools</font></td>	
		<td><font size=1 face="ms sans serif, arial, geneva">
<% for i = 1 to 4 %>
			<% =i %>.&nbsp;&nbsp;&nbsp;<% =rs("College") %>&nbsp;&nbsp;&nbsp;<font color=gray>(<% =rs("Cnt") %>)</font><br>
<%
	rs.movenext
next
%>
			</font></td>
		</tr>
<% rs.close %>



	<tr><td colspan=3><font size=1 face="ms sans serif, arial, geneva">
		<br><b>Technology Statistics</b>
		</font></td>



<%
sql =	"SELECT Tech.Tech_ID, Tech.Tech_Name, " &_
		"    COUNT(Employee_Tech_Xref.Employee_Tech_Xref_ID) " &_
		"    AS cnt " &_
		"FROM Tech INNER JOIN " &_
		"    Employee_Tech_Xref ON " &_
		"    Tech.Tech_ID = Employee_Tech_Xref.Tech_ID INNER JOIN " &_
		"    employee ON " &_
		"    Employee_Tech_Xref.employee_id = employee.employeeID " &_
		"GROUP BY Tech.Tech_ID, Tech.Tech_Name " &_
		"ORDER BY cnt DESC"
set rs = DBQuery(sql)
%>
	<tr><td width=10></td><td valign=top><font size=1 face="ms sans serif, arial, geneva">
		Sark Expertise</font></td>	
		<td><font size=1 face="ms sans serif, arial, geneva">
<% for i = 1 to 8 %>
			<% =i %>.&nbsp;&nbsp;&nbsp;<% =rs("Tech_Name") %>&nbsp;&nbsp;&nbsp;<font color=gray>(<% =rs("Cnt") %>)</font><br>
<%
	rs.movenext
next
%>
			</font></td>
		</tr>
<% rs.close %>



</table>

</td></tr></table>
</center>


<!-- #include file="../../footer.asp" -->
