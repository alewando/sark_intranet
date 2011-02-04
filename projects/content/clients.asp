<!-- #include file="../section.asp" -->


<%
'---------------------------------------------------------------------------------
' Description:	List of projects by client.
' History:		06/02/1999 - KDILL - Created
'---------------------------------------------------------------------------------
%>


<script language=javascript>
<!--
function AddProject(clientid){
	winTech = window.open("details_add.asp?ClientID=" + clientid, "AddProject", "height=350, width=600, toolbar=0, location=0, directories=0, status=0, menubar=0, resizable=1, scrollbar=1")
	}
function EditProject(clientid, projectid){
	winTech = window.open("details_add.asp?ClientID=" + clientid + "&ProjectID=" + projectid, "AddProject", "height=350, width=600, toolbar=0, location=0, directories=0, status=0, menubar=0, resizable=1, scrollbar=1")
	}
// -->
</script>


To view the complete details of a client, including all project information,
click on the client's name.  To view only a specific project, click on the
project name.<p>

<%
AddProject = Session("OfficeStaff")
EditProject = Session("OfficeStaff")
sql =	"SELECT       Client.Client_ID, Client.ClientName, Project.Project_ID, Project.Project_Name, Project.Client_Contacts, Project.StartDate, Project.EndDate, Project.Project_Desc " & _
		"FROM         Project RIGHT OUTER JOIN Client ON Project.Client_ID = Client.Client_ID " & _
		"WHERE        Client.SortOverride = 0 AND Client.Client_ID > 1" & _
		"ORDER BY Client.ClientName, Project.Project_Name"
set rs = DBQuery(sql)
response.write("<!-- sql = " & sql & " -->" & NL)
while not rs.eof
	RecordsetMoved = false
	ClientID = rs("Client_ID")
%>

<table class=tableShadow width="100%" cellspacing=0 cellpadding=4 border=0 bgcolor=#ffffcc>
	<tr>
		<td><font size=1 face="ms sans serif, arial, geneva">
			<b><a href="details.asp?clientid=<%=rs("Client_ID")%>&"><%=rs("ClientName")%></a></b>
			</font></td>
		<td align=right><font size=1 face="ms sans serif, arial, geneva" color=black>
<%if AddProject then%>
			[<a href="javascript:AddProject(<% =ClientID %>);" onMouseOver="top.status='Add new project.'; return true;" onMouseOut="top.status=''; return true;">Add</a>]
<%end if%>
		</FONT></td>
	</tr>

<%
if (not rs.eof) and (not isnull(rs("Project_Name"))) then response.write("	<tr><td colspan=2 bgcolor=gray height=1></td></tr>")

continue = (not isnull(rs("Project_Name"))) and (ClientID = rs("Client_ID"))
while (not rs.eof) and continue
	RecordsetMoved = true
	proj_desc = rs("Project_Desc")
	proj_desc = trim(left(proj_desc, 170))
	i = InStrRev(proj_desc, " ")
	if i > 0 then proj_desc = left(proj_desc, i-1) & "..."
%>

	<tr>
		<td valign=top><font size=1 face="ms sans serif, arial, geneva" color=black>
		<a href="details.asp?clientid=<% =rs("Client_ID") %>&projectid=<% =rs("Project_ID") %>">
		<b><%=rs("Project_Name")%></b></a>
<%
if not isnull(rs("StartDate")) then
	response.write("<font color=black>&nbsp;&nbsp;(" & Month(rs("StartDate")) & "/" & Year(rs("StartDate")) & " - ")
	if not isnull(rs("EndDate"))then
		response.write(Month(rs("EndDate")) & "/" & Year(rs("EndDate")))
	else
		response.write("Present")
		end if
	response.write(")</font>")
	end if
%>
		<br><% =proj_desc %>
		</font></td>

<%if EditProject then%>
		<td align=right valign=top nowrap><font size=1 face="ms sans serif, arial, geneva" color=black>
		[<a href="javascript: EditProject(<% =rs("Client_ID") %>, <%=rs("Project_ID")%>)" onMouseOver="top.status='Edit project.';return true;" onMouseOut="top.status='';return true;">Edit</a>]
		</font></td>
<%else%>
		<td>&nbsp;</td>
<%end if%>
	</tr>

<%
	rs.movenext
	if rs.eof then
		continue = false
	else
		continue = (not isnull(rs("Project_Name"))) and (ClientID = rs("Client_ID"))
		end if
	wend
%>

</table><p>

<%
	if (not rs.eof) and (not RecordsetMoved) then rs.movenext
	wend
rs.close
set rs = nothing
%>


<!-- #include file="../../footer.asp" -->
