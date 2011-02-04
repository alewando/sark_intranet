<!--
Developer:    Adam Lewandowski
Date:         10/27/2000
Modified:	  12/26/2000
Modified By:  Julie Walters
Description:  Part of the amReports group.  Lists the security roles of employees.
-->
<!-- #include file="../section.asp" -->
<%if (hasRole("WebMaster")) then %>

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
<b>Security roles assigned to Employees</b>
<p>
<font size=2 face="ms sans serif, arial, geneva"  color=red>
<b>Click on a hyperlink below to go directly to a Sark profile.</b>
<p>

<!-- Start Table 1-->
<TABLE BORDER=0 WIDTH=450>
<tr><td ALIGN=LEFT VALIGN=top>
<%
sql= "SELECT ur.employee_id, ur.role_id, r.role_description, e.FirstName, e.LastName " & _
     "FROM Security_Roles r, Security_User_Roles ur, Employee e " & _
     "WHERE r.role_id = ur.role_id AND ur.employee_id = e.employeeid " & _
     "ORDER BY r.role_id, LastName, FirstName"
set rs = DBQuery(sql)
%>
<!--Start Table 1.1-->
<table width=375 cellspacing=0 cellpadding=4 border=0 >
<tr><td colspan=2 height=1>
</td></tr>

<%
role = "" 
while not rs.eof
	
	id = (rs("Employee_ID"))
	roleID = rs("role_id")
	fullName = rs("LastName")&", "& rs("FirstName")
	roleName = rs("role_description")
	role = roleName
	if role = "" OR role <> lastRole then%>
	<tr>
		<td><font size=1 face= "ms sans serif, arial, geneva"  color = blue>
		<b><span onClick="toggle('<%=roleName%>')" alt="Expand or collapse this section." style="cursor:hand"><%=roleName%></b>
		<span id='<%=roleName & "Header"%>' style="">(collapsed)</span>
		</td>	</tr>	<% lastRole = roleName	end if	%>		<tr><td colspan=2>
		<span id='<%=roleName%>' style="display:none"><!--Start Table 1.1.1-->
			<table>			<%			while role = roleName %>
				<tr>
				 <td valign=middle align=left>
				  <font size=1 face='ms sans serif, arial, geneva'>				  				  <A HREF="../../directory/content/details.asp?EmpID=<%=id%>"><%=fullName%></A>
				  </font>				 </td>				</tr>
			 				<%	rs.movenext									if not rs.eof then					id = (rs("Employee_ID"))
					roleID = rs("role_id")
					fullName = rs("LastName")&", "& rs("FirstName")
					roleName = rs("role_description")
				else					roleName=""
				end if			wend%>	<!--Close Table 1.1.1-->
			</table>
		</span>		
	</td></tr>	
<%wend
rs.close%><!--Close Table 1.1-->
</table>

</td></tr>
<!--Add Sports team managers listing in next rows-->
<tr></tr><tr><td ALIGN=LEFT VALIGN=top>
<%
sql= "SELECT sm.employee_id, s.sport_id, s.sport_name, e.FirstName, e.LastName " & _
     "FROM Sports_managers sm, sports s, Employee e " & _
     "WHERE s.sport_id = sm.sport_id AND sm.employee_id = e.employeeid " & _
     "ORDER BY s.sport_id, e.LastName, e.FirstName"
set rs = DBQuery(sql)
%>
<b>Sports Managers</b>
<!--Start Table 1.2-->
<table width=375 cellspacing=0 cellpadding=4 border=0 >
<tr><td colspan=2 height=1>
</td></tr>

<%
role = "" 
while not rs.eof
	
	id = (rs("Employee_ID"))
	roleID = rs("sport_id")
	fullName = rs("LastName")&", "& rs("FirstName")
	roleName = rs("Sport_name")
	role = roleName
	if role = "" OR role <> lastRole then%>
	<tr>
		<td><font size=1 face= "ms sans serif, arial, geneva"  color = blue>
		<b><span onClick="toggle('<%=roleName%>')" alt="Expand or collapse this section." style="cursor:hand"><%=roleName%></b>
		<span id='<%=roleName & "Header"%>' style="">(collapsed)</span>
		</td>	</tr>	<% lastRole = roleName	end if	%>		<tr><td colspan=2>
		<span id='<%=roleName%>' style="display:none"><!--Start Table 1.2.1-->
			<table>			<%			while role = roleName %>
				<tr>
				 <td valign=middle align=left>
				  <font size=1 face='ms sans serif, arial, geneva'>				  				  <A HREF="../../directory/content/details.asp?EmpID=<%=id%>"><%=fullName%></A>
				  </font>				 </td>				</tr>
			 				<%	rs.movenext									if not rs.eof then					id = (rs("Employee_ID"))
					roleID = rs("sport_id")
					fullName = rs("LastName")&", "& rs("FirstName")
					roleName = rs("sport_name")
				else					roleName=""
				end if			wend%>	<!--Close Table 1.2.1-->
			</table>
		</span>		
	</td></tr>	
<%wend
rs.close%><!--Close Table 1.2-->
</table>
</td>
</tr><!--Close Table 1-->
</table>
	
<!-- #include file="../../footer.asp" -->
<% else %>
	<b><h2><font color=red>You are not currently a webmaster<br>, therefore can not view this report.
	</b></h2></font>
<%end if%>

