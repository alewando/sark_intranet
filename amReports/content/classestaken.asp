<!--
Developer:   DMARTIN
Date:         10/2000

			
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
<b>Consultants listed by Classes Taken</b>
<br></br>
<font size=2 face="ms sans serif, arial, geneva"  color=red>
<b>Click on a hyperlink below to go directly to a Sark profile.</b><p>

<TABLE BORDER=0 WIDTH=450>
<tr><td ALIGN=LEFT VALIGN=top>
<!--''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''-->
<%
sql= "select employee.EmployeeID, employee.firstname, employee.lastname, " &_
     "classes.class_Start_date, classes.class_End_date, classes.class_id, classes.employeeid, class_list.class_id, class_list.class_name " &_
     "FROM class_list LEFT OUTER JOIN Classes INNER JOIN Employee ON " &_
     "Classes.employeeid = Employee.employeeid ON " &_
     "Class_list.class_id = classes.class_id " &_
     "ORDER BY  class_list.class_name"
     
     
     
     
set rs = DBQuery(sql)

name =""
%>
<table width=490 cellspacing=0 cellpadding=4 border=0 >
<tr><td colspan=2 height=1>
</td></tr>

<%while not rs.eof
	
	id = (rs("EmployeeID"))
	classid = rs("class_id")
	first_name = rs("FirstName")
	last_name = rs("LastName")
	classn = rs("class_name")
	taken = rs("class_Start_date")
	
	if name = "" OR name <> classn then%>
	<tr>
		<td><font size=1 face= "ms sans serif, arial, geneva"  color = blue>
		<b><span onClick="toggle('<%=classn%>')" alt="Expand or collapse this section." style="cursor:hand"><%%>&nbsp;&nbsp;<%=classn%></b>
		<span id='<%=classn & "Header"%>' style="">(collapsed)</span>
		<!--<b><%=title%></b>-->
		</td>
	<%end if%>
		<span id='<%=classn%>' style="display:none">
				
				Response.write("<td colspan=1 width=350valign=middle align=left><font size=1 face='ms sans serif, arial, geneva'>")
				Response.Write("<b>Consultant:</b>")
				Response.write("<td colspan=1 width=80 valign=middle align=left><font size=1 face='ms sans serif, arial, geneva'>")
				Response.Write("<b>Date Passed: </b>")%>		
				<tr>
				
				<td colspan=1 width=350 valign=top><font size=1 face= "ms sans serif, arial, geneva" color=black>
			 
						id = (rs("EmployeeID"))
						first_name = rs("FirstName")
						last_name = rs("LastName")
						classn = trim(rs("class_name"))
					else
					end if
						
			</table>
		</span>
	</td></tr>
	

<%wend
rs.close%>

</tr>

</table>
	
<!-- #include file="../../footer.asp" -->
<% else %>
	<b><h2><font color=red>You are not currently an account manager or sales representative,<br> therefore can not view the Account Manager Reports.
	</b></h2></font>
<%end if%>
