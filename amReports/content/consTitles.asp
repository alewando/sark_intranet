<!--
Developer:    SSeissiger
Date:         8/2000

			
-->
<%if (hasRole("AccountManager")) then %>
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
<b>Consultants listed by Title</b>
<br></br>
<font size=2 face="ms sans serif, arial, geneva"  color=red>
<b>Click on a hyperlink below to go directly to a Sark profile.</b>
<p>
<TABLE BORDER=0 WIDTH=450>
<tr><td ALIGN=LEFT VALIGN=top>
<!--''''''''''''''''''''''''''''''''''''''''''''''
'Code to display a list of the Industry 
'Experts sorted alphabetically by Industry by
'consultant.
''''''''''''''''''''''''''''''''''''''''''''''-->
<%
sql= "select e.EmployeeID,e.firstname, e.lastname, t.title_desc " &_
"from employee e, employee_title t " &_
"where e.employeetitle_id = t.employee_title_id " &_
"ORDER BY t.employee_title_id, e.lastname"
'"and t.title_desc in ('Consultant', 'Senior Consultant', 'Managing Consultant', " &_
'"'Project Manager','Technical Manager','Solution Services Practice Lead')" &_
     
set rs = DBQuery(sql)

prev_title = ""

%>
<table width=375 cellspacing=0 cellpadding=4 border=0 >
<tr><td colspan=2 height=1>
</td></tr>
<%while not rs.eof
	
	id = (rs("EmployeeID"))
	first_name = rs("FirstName")
	last_name = rs("LastName")
	title = trim(rs("title_desc"))

	if prev_title = "" OR prev_title <> title then%>
	<tr>
		<td><font size=1 face= "ms sans serif, arial, geneva"  color = maroon>
		<b><span onClick="toggle('<%=title%>')" alt="Expand or collapse this section." style="cursor:hand"><%=title%></b>
		<span id='<%=title & "Header"%>' style="">(collapsed)</span>
		</td>	</tr>
	<%end if%>	<%prev_title = title%>		<tr><td colspan=2>
		<span id='<%=title%>' style="display:none">			<table>									<% while prev_title = title%>
				<tr>
				<td valign=top><font size=1 face= "ms sans serif, arial, geneva" color=black>				&nbsp;&nbsp;<a href='../../directory/content/details.asp?EmpID=<%=id%>'><%=first_name%>&nbsp;<%=last_name%></a><br>				</td></tr>
			 				<%	rs.movenext										if not rs.eof then
						id = (rs("EmployeeID"))
						first_name = rs("FirstName")
						last_name = rs("LastName")
						title = trim(rs("title_desc"))
					else						prev_title=""
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
	<b><h2><font color=red>You are not currently an account manager,therefore can not view the Account Manager Reports.
	</b></h2></font>

<%end if%>

