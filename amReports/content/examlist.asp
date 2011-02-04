<!--
Developer:   DMARTIN
Date:         9/2000

			
-->
<%if (hasRole("AccountManager") or hasRole("Sales")) then %>
<!-- #include file="../section.asp" -->

<script language=javascript>
function toggle(item){
	var docItem = document.all(item)
	if (docItem.style.display == ""){
		docItem.style.display = "none"
		document.all(item + "Header").innerHTML = "(collapsed)"
		}
	else {
		docItem.style.display = ""
		document.all(item + "Header").innerHTML = ""
		}
	}
</script>
<font size=3 face='ms sans serif, arial, geneva' color=black>
<b>Consultants listed by Exams Passed</b>
<p>
<font size=2 face="ms sans serif, arial, geneva"  color=red>
<b>Click on a hyperlink below to go directly to a Sark profile.</b><p>

<TABLE BORDER=0 WIDTH=450>
<tr><td ALIGN=LEFT VALIGN=top>
<!--''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''-->
<%
sql= "select employee.EmployeeID, employee.firstname, employee.lastname, exam_list.cert, " &_
     "exams_passed.exam_date, exams_passed.exam_id, exams_passed.employeeid, exam_list.exam_id, exam_list.exam_name " &_
     "FROM exam_list LEFT OUTER JOIN Exams_passed INNER JOIN Employee ON " &_
     "Exams_passed.employeeid = Employee.employeeid ON " &_
     "Exam_list.exam_id = Exams_passed.exam_id " &_
     "WHERE Exam_list.cert = 0 " &_
     "ORDER BY  exam_list.exam_name"
     
     
     
     
set rs = DBQuery(sql)

name =""
%>
<table width=375 cellspacing=0 cellpadding=4 border=0 >
<tr><td colspan=2 height=1>
</td></tr>

<%while not rs.eof
	
	id = (rs("EmployeeID"))
	examid = rs("exam_id")
	first_name = rs("FirstName")
	last_name = rs("LastName")
	exam = rs("exam_name")
	passed = rs("exam_date")
	
	if name = "" OR name <> exam then%>
	<tr>
		<td><font size=1 face= "ms sans serif, arial, geneva"  color = blue>
		<b><span onClick="toggle('<%=exam%>')" alt="Expand or collapse this section." style="cursor:hand"><%%>&nbsp;&nbsp;<%=exam%></b>
		<span id='<%=exam & "Header"%>' style="">(collapsed)</span>
		<!--<b><%=title%></b>-->
		</td>	</tr>	
	<%end if%>	<%name = exam%>		<tr><td colspan=2>
		<span id='<%=exam%>' style="display:none">			<table>				<%				Response.write("<table><tr><font size=1 face='ms sans serif, arial, geneva'>")
				
				Response.write("<td colspan=1 width=350valign=middle align=left><font size=1 face='ms sans serif, arial, geneva'>")
				Response.Write("<b>Consultant:</b>")
				Response.write("<td colspan=1 width=80 valign=middle align=left><font size=1 face='ms sans serif, arial, geneva'>")
				Response.Write("<b>Date Passed: </b>")%>					<% while name = exam%>
				<tr>
				
				<td colspan=1 width=350 valign=top><font size=1 face= "ms sans serif, arial, geneva" color=black>				<a href='../../directory/content/details.asp?EmpID=<%=id%>'><%=first_name%>&nbsp;<%=last_name%></a><br>				<td colspan=1 width=80 valign=middle align=left><font size=1 face='ms sans serif, arial, geneva'><%=passed%>				</td></tr>
			 				<%	rs.movenext										if not rs.eof then
						id = (rs("EmployeeID"))
						first_name = rs("FirstName")
						last_name = rs("LastName")
						exam = trim(rs("exam_name"))						passed = rs("exam_date")
					else						name=""
					end if			   wend%>	
						
			</table>
		</span>		
	</td></tr>
	

<%wend
rs.close%>
</td>
</tr></table>
</tr>
</table>
	
<!-- #include file="../../footer.asp" -->
<% else %>
	<b><h2><font color=red>You are not currently an account manager or sales representative,<br> therefore can not view the Account Manager Reports.
	</b></h2></font>
	
<%end if%>

