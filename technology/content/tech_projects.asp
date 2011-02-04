<%
'---------------------------------------------------------------------------------
' Description:	Include file for technology section to display project listings.
' History:		06/03/1999 - KDILL - Created
'---------------------------------------------------------------------------------
%>


<script language=javascript>
<!--
function EditProject(tech_links_id){
	winTech=window.open("links_edit.asp?empname=<%=Session("Username")%>&TechName=<%=Server.URLEncode(curTitle)%>&Tech_ID=<%=techid%>&Recipients=<%=Recipients%>&Link_ID=" + tech_links_id, "EditLink","height=250,width=580,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1, scrollbar=1")
	}
// -->
</script>


<table width="100%" cellspacing=0 cellpadding=4 border=0 bgcolor=#ffffcc class=tableShadow>
	<tr>
		<td><font size=1 face="ms sans serif, arial, geneva" color=navy><b>
			Sark Projects
			</b></font></td>
			
		<td align=right><font size=1 face="ms sans serif, arial, geneva" color=black>
			[<a href="../../projects/content/clients.asp" onMouseOver="top.status='View project listings.'; return true;" onMouseOut="top.status=''; return true;">Summary</a>]</font></td>

		</tr>
		
	<tr><td colspan=2 bgcolor=gray height=1></td></tr>

	
<%
sql =	"SELECT Project.Project_ID, Project.Project_Name, Client.Client_ID, Client.ClientName, Project.Project_Desc " & _
		"FROM Project INNER JOIN Project_Tech_Xref ON Project.Project_ID = Project_Tech_Xref.Project_ID INNER JOIN Client ON Project.Client_ID = Client.Client_ID " & _
		"WHERE (Project_Tech_Xref.Tech_ID = " & techid & ") " & _
		"ORDER BY Project.StartDate DESC, ClientName, Project_Name"
set rs = DBQuery(sql)
if rs.eof then
	response.write("<tr><td colspan=2><font size=1 color=gray face='ms sans serif, arial, geneva'>[None identified]</td></tr>")
else
	while not rs.eof
		set textobj = rs("Project_Desc")
		textval = trim(left(textobj, 100))
		i = InStrRev(textval, " ")
		if i > 0 then textval = left(textval, i-1) & "..."
		if len(textval) > 0 then textval = "<br>" & textval
%>

	<tr>
		<td colspan=2><font size=1 face="ms sans serif, arial, geneva" color=black>
			<b><a href="../../projects/content/details.asp?clientid=<% =rs("Client_ID") %>&"><% =rs("ClientName") %></a>&nbsp;-&nbsp;</b>
			<b><a href="../../projects/content/details.asp?clientid=<% =rs("Client_ID") %>&projectid=<% =rs("Project_ID") %>"><% =rs("Project_Name") %></a></b><% =textval %><br>
			<font color=navy>&nbsp;&nbsp;-&nbsp;&nbsp;
<%
sql =	"SELECT      Employee.Username, Employee.EmployeeID, Employee.FirstName, Employee.LastName " & _
		"FROM         Employee_Project_Xref INNER JOIN Employee ON Employee_Project_Xref.Employee_ID = Employee.EmployeeID " & _
		"WHERE      Employee_Project_Xref.Project_ID = " & rs("Project_ID") & " " & _
		"ORDER BY   Employee.LastName"
set rs2 = DBQuery(sql)
if not rs2.eof then
	while not rs2.eof
		response.write("<a href='../../directory/content/details.asp?EmpID=" & rs2("EmployeeID") & "'>" & rs2("FirstName") & " " & rs2("LastName") & "</a>")
		rs2.movenext
		if not rs2.eof then response.write(", ")
		wend
	rs2.close
	end if
%>
			</font><br>
		</font></td>

	</tr>

<%
		rs.movenext
		wend
	rs.close
	set rs = nothing
end if
%>


</table>
