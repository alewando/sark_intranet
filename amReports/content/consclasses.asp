<!--
Developer:    DMARTIN
Date:         9/2000

			
-->
<%'if (hasRole("AccountManager")) then %>
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
<b>Classes Listed by Consultants who have taken classes</b>
<p>
<font size=2 face="ms sans serif, arial, geneva"  color=red>
<b>Click on a hyperlink below to go directly to a Sark profile.</b>
<p>

<TABLE BORDER=0 WIDTH=450>
<tr><td ALIGN=LEFT VALIGN=top>
<!--''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''-->
<%
   
      
sql= "select Employee.EmployeeID, Employee.firstname, Employee.lastname, " &_
	 "Classes.class_Start_date, Classes.class_End_date, Classes.class_id, Classes.employeeid, Class_List.class_id, Class_List.class_name " &_
	 "FROM Employee LEFT OUTER JOIN Classes INNER JOIN Class_List ON class_list.class_id = classes.class_id ON " &_
	 "employee.employeeid = classes.employeeid " &_
	 "WHERE Employee.EmployeeTitle_ID in (1, 2, 3, 4, 5, 6, 7, 8, 23) " &_
	 "ORDER BY Employee.lastname, Employee.firstname"
      
      
      
set rs = DBQuery(sql)

name =""
%>
<table width=375 cellspacing=0 cellpadding=4 border=0 >
<tr><td colspan=2 height=1>
</td></tr>

<%while not rs.eof
	
	id = (rs("EmployeeID"))
	classid = rs("class_id")
	full_name = rs("LastName")&", "& rs("FirstName")
	classname = rs("class_name")
	taken = rs("class_Start_date")
	
	if name = "" OR name <> lastname then%>
	<tr>
		<td><font size=1 face= "ms sans serif, arial, geneva"  color = blue>
		<b><span onClick="toggle('<%=full_name%>')" alt="Expand or collapse this section." style="cursor:hand"><%=full_name%></b>
		<span id='<%=full_name & "Header"%>' style="">(collapsed)</span>
		<!--<b><%=title%></b>-->
		</td>	</tr>	
	<%end if%>	<%name = full_name%>		<tr><td colspan=2>
		<span id='<%=full_name%>' style="display:none">			<table>				<%				Response.write("<table><tr><font size=1 face='ms sans serif, arial, geneva'>")
				Response.write("<td colspan=1 width=340 valign=middle align=left><font size=1 face='ms sans serif, arial, geneva'>")
				Response.Write("<b>Classes Taken:</b>")
				Response.write("<td colspan=1 width=40 valign=middle align=left><font size=1 face='ms sans serif, arial, geneva'>")
				Response.Write("<b>Date Taken: </b>")%>					<% while name = full_name %>
				<tr>
				
				<td colspan=1 width=340 valign=top><font size=1 face= "ms sans serif, arial, geneva" color=black>				<%=classname%></a><br>				<td colspan=1 width=40 valign=middle align=left><font size=1 face='ms sans serif, arial, geneva'><%=taken%>				</td></tr>
			 				<%	rs.movenext										if not rs.eof then
						id = (rs("EmployeeID"))
						classid = rs("class_id")
						full_name = rs("LastName")&", "& rs("FirstName")
						classname = rs("class_name")						taken = rs("class_Start_date")
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
<% 'else %>
	<!--<b><h2><font color=red>You are not currently an account manager,<br> therefore can not view the Account Manager Reports.
	</b></h2></font>
	-->
<%'end if%>

