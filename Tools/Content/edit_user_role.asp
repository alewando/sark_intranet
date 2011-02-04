<%@ Language = Vbscript%>
<!--
Developer:	  Adam Lewandowski
Date:         10/16/2000
Description:  Allows the webmaster to assign users to roles
-->

<!-- #include file="../section.asp" -->

<HTML>

<HEAD>

</HEAD>



<%
		employee_ID=Request("employee_ID")
		if IsNull(employee_ID) or employee_ID = "" then
			employee_ID = "-1"
		end if
		newRoleID=Request("newRoleID")
		deletedRoleID=Request("deletedRoleID")

		if Request.Form("Submitted")="True" then
			' Add the role
			sql = "INSERT INTO Security_User_Roles (employee_id, role_id) VALUES " & _
			      "(" & employee_id & ", " & newRoleID & ")"
			DBQuery(sql)		
		elseif Request.Form("deleted")="True" then
			' Delete the role
			sql = "DELETE FROM Security_User_Roles WHERE " & _
			      "employee_id=" & employee_id & " AND role_id=" & deletedRoleID
			DBQuery(sql)				
		elseif Request.form("Selected")="True" then
			' Update the employee information
		end if
	
%>
<body>

<table border=0><tr><td><font size=1 face="ms sans serif, arial, geneva" color=blue><b><STRONG>
  <font size=2 color=red><b>Edit User Roles</b></font><p>
  <% if not hasRole("WebMaster") then
	Response.Write("<B>Only the Webmaster can assign roles</B><P>")
  end if %>
  </font></td></tr></table>

<% if hasRole("WebMaster") then %>
<form NAME="frmInfo" ACTION="<%=Request.ServerVariables("SCRIPT_NAME")%>" method=post>
  <INPUT TYPE=Hidden NAME="InsertUpdate" VALUE="<%=ls_empID%>"><br>
  <INPUT TYPE=HIDDEN NAME="Submitted" VALUE="False">
  <INPUT TYPE=Hidden NAME="Selected" VALUE="False">
  <INPUT TYPE=Hidden NAME="deleted" VALUE="False">
  <INPUT TYPE=Hidden NAME="deletedRoleID" VALUE="-1">

  <table border=0>
  <tr>
	<td align=left><font face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Employee:</STRONG></font></td>
	<td>
		<Select NAME="employee_ID" onChange='javascript:LoadForm();'>
			<%	set rs = DBQuery("select employeeid, firstName, lastName from employee ORDER BY lastname, firstname")
				if not rs.eof and not rs.bof then
					do while not rs.eof
						Response.Write("<Option Value=" & trim(rs("employeeid")))
						if cInt(trim(rs("employeeid")))=cInt(employee_ID) then
							Response.Write(" SELECTED ")
						end if
						Response.Write(">" & rs("lastName") & ", " & rs("firstName") & vbCRLF) 
						rs.movenext
					loop
					rs.close
				end if
				%>
		</Select>
	</td>
  </tr>

  <tr><td align=left><font face="ms sans serif, arial, geneva" size=1 color=blue><STRONG>Roles:</STRONG></font></td>
	<td></td></tr>
	
	<%
	sql = "SELECT r.role_id, r.role_name, r.role_description FROM Security_Roles r, Security_User_Roles ur " & _
	      "WHERE ur.employee_id=" & employee_ID & " AND r.role_id = ur.role_id"
	set rsRoles = DBQuery(sql)
	while not rsRoles.eof
	 Response.Write("<TR><TD></TD><TD>")
	 Response.Write(rsRoles("role_description"))
	 Response.Write("</TD><TD><INPUT TYPE=BUTTON onClick='javascript:deleteRole(" & rsRoles("role_id") & ")' VALUE='Remove Role' class=button>")
	 Response.Write("</TD></TR>")
	 rsRoles.moveNext
	wend
	%>
	<TR><TD></TD><TD>
	<SELECT NAME="newRoleID">
	<%
	sql = "SELECT role_id, role_description FROM Security_Roles WHERE role_id not in (SELECT role_id FROM Security_User_Roles WHERE employee_id=" & employee_ID & ") ORDER BY role_name"
	set rsRoles = DBQuery(sql)
	while not rsRoles.eof
	 Response.Write("<OPTION VALUE='" & rsRoles("role_id") & "'>" & rsRoles("role_description") & vbCRLF)
	 rsRoles.moveNext
	wend
	%>
	</SELECT></TD>
	<TD><INPUT TYPE=BUTTON class=button VALUE="Add Role" onClick='javascript:addRole()'></TD>
	</TR>
</table>


</form>
</body>
<SCRIPT LANGUAGE=javascript>
<!--

	function GoHome()
	{
		window.navigate("default.asp")
	}
	function LoadForm()
	{
		document.frmInfo.Selected.value="True"
		document.frmInfo.submit()
	}
	
	function addRole() {
		document.frmInfo.Submitted.value="True"
		document.frmInfo.submit()
	}
	
	function deleteRole(roleID)
	{
 		if (confirm("Are you sure you want to remove the user from this role?")) {
			document.frmInfo.deleted.value="True"
			document.frmInfo.deletedRoleID.value=roleID
					document.frmInfo.submit()
		}
	}
	
//-->
</SCRIPT>

<% end if ' End If hasRole("WebMaster") %>
</HTML>







<!-- #include file="../../footer.asp" -->