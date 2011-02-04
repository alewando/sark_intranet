<!-- #include file="../section.asp" -->


<%
'---------------------------------------------------------------------------------
' Description:	Display project details.
' History:		06/03/1999 - KDILL - Created
'---------------------------------------------------------------------------------
%>


<script language=javascript>
<!--
function EditClient(){
	winTech = window.open("details_client.asp?ClientID=<% =request("ClientID") %>", "EditClient", "height=350, width=600, toolbar=0, location=0, directories=0, status=0, menubar=0, resizable=1, scrollbar=1")
	}
function AddProject(){
	winTech = window.open("details_add.asp?ClientID=<% =request("ClientID") %>", "AddProject", "height=350, width=600, toolbar=0, location=0, directories=0, status=0, menubar=0, resizable=1, scrollbar=1")
	}
function EditProject(ProjectID){
	winTech = window.open("details_add.asp?ClientID=<% =request("ClientID") %>&ProjectID=" + ProjectID, "AddProject", "height=350, width=600, toolbar=0, location=0, directories=0, status=0, menubar=0, resizable=1, scrollbar=1")
	}
function AddEmps(ClientID, ProjectID, ProjectName){
	winTech = window.open("details_add_emps.asp?ClientID=" + ClientID + "&ProjectID=" + ProjectID + "&ProjectName=" + ProjectName, "AddEmps", "height=400, width=250, toolbar=0, location=0, directories=0, status=0, menubar=0, resizable=1, scrollbars=1")
	}
function AddTechs(ClientID, ProjectID, ProjectName){
	winTech = window.open("details_add_tech.asp?ClientID=" + ClientID + "&ProjectID=" + ProjectID + "&ProjectName=" + ProjectName, "AddTech", "height=400, width=350, toolbar=0, location=0, directories=0, status=0, menubar=0, resizable=1, scrollbars=1")
	}
// -->
</script>


<%
if Session("OfficeStaff") then
	EditClient = true
	AddProjects = true
	EditProject = true
	end if

sql = "SELECT Client_ID, ClientName, Description, WebSite FROM Client WHERE (Client.SortOverride = 0) "
if request("ClientID") <> "" then sql = sql & "AND (Client.Client_ID = " & request("ClientID") & ") "
sql = sql & "ORDER BY Client.ClientName"
set rs = DBQuery(sql)
while not rs.eof
	desc = rs("Description")
	if desc <> null then desc = Replace(desc, chr(13), "<br>")
	response.write("<font size=2><b>" & rs("ClientName") & "</b></font><br>")
%>

<table width="100%" border=0 cellpadding=0 cellspacing=0><tr>
	<td valign=bottom><font size=1 face="ms sans serif, arial, geneva" color=black>
		<% =desc%><% if not isnull(rs("WebSite")) then response.write("&nbsp;&nbsp;[<a href='http://" & rs("WebSite") & "' target='_blank'>Website</a>]") %>
		</font></td>
	<td width=10>&nbsp;</td>
	<td align=right valign=bottom nowrap>
<%if EditClient then%>
		<font size=1 face="ms sans serif, arial, geneva">
			[<a href="javascript:EditClient();" onMouseOver="top.status='Edit client Information.'; return true;" onMouseOut="top.status=''; return true;">Edit</a>]
		</FONT>
<%end if%>
	</td>
</tr></table><br>


<table class=tableShadow width="100%" cellspacing=0 cellpadding=4 border=0 bgcolor=#ffffcc>
	<tr>
		<td colspan=4><font size=1 face="ms sans serif, arial, geneva" color=navy>
			<b>Projects</b>
<% if request("projectid") = "" then response.write("<font color=gray>&nbsp;&nbsp;(listed by start date and project name)</font>") %>
			</font></td>

		<td align=right valign=top nowrap><font size=1 face="ms sans serif, arial, geneva" color=black>
<% if AddProjects then %>
		[<a href="javascript:AddProject();" onMouseOver="top.status='Add new project.'; return true;" onMouseOut="top.status=''; return true;">Add</a>]
<% end if %>
<% if request("projectid") <> "" then %>
			[<a href="details.asp?clientid=<% =request("clientid") %>&" onMouseOver="top.status='View all project listings.';return true;" onMouseOut="top.status='';return true;">View All Projects</a>]
<% end if %>
			</font></td>
		</tr>

	<tr><td colspan=5 bgcolor=gray height=1></td></tr>

<%
sql =	"SELECT Project_ID, Project_Name, Client_Contacts, OtherTech, StartDate, EndDate, Project_Desc " & _
		"FROM Project " & _
		"WHERE (Client_ID = " & rs("Client_ID") & ") "
if request("projectid") <> "" then sql = sql & " AND (Project_ID = " & request("projectid") & ") "
sql = sql & "ORDER BY StartDate DESC, Project_Name"
set rs2 = DBQuery(sql)
if rs2.eof then
	response.write("<tr><td colspan=4><font size=1 face='ms sans serif, arial, geneva' color=black>")
	response.write("<font color=gray>[None identified]</font>")
	response.write("</td></tr>")
else
	while not rs2.eof
%>

	<tr><td colspan=4><font size=1 face="ms sans serif, arial, geneva" color=black>
		<b><a href="details.asp?clientid=<% =request("ClientID") %>&projectid=<% =rs2("Project_ID") %>&"><% =rs2("Project_Name") %></a></b>
<%
if not isnull(rs2("StartDate")) then
	response.write("&nbsp;&nbsp;(" & Month(rs2("StartDate")) & "/" & Year(rs2("StartDate")) & " - ")
	if not isnull(rs2("EndDate")) then
		response.write(Month(rs2("EndDate")) & "/" & Year(rs2("EndDate")))
	else
		response.write("Present")
		end if
	response.write(")")
	end if
 %>
		</td>
		<td align=right valign=top><font size=1 face="ms sans serif, arial, geneva" color=black>
<% if EditProject then %>
		[<a href="javascript: EditProject(<%=rs2("Project_ID")%>)" onMouseOver="top.status='Edit or delete project.';return true;" onMouseOut="top.status='';return true;">Edit</a>]
<% else %>
		&nbsp;
<% end if %>
		</font></td></tr>

	<tr>
		<td colspan=5><font size=1 face="ms sans serif, arial, geneva" color=black>
<%
proj_desc = rs2("Project_Desc")
if proj_desc <> null then proj_desc = Replace(proj_desc, chr(13), "<br>")
response.write(proj_desc)
if rs2("Client_Contacts") <> "" then
	response.write(" <font color=gray>(" & rs2("Client_Contacts") & ")</font>")
	end if
%>
			</font></td>
	</tr>
		
	<tr>
		<td width=10% align=left valign=top><font size=1 face="ms sans serif, arial, geneva" color=black>
<% if AddProjects then%>
			[<a href="javascript: AddEmps(<% =request("ClientID") %>, <% =rs2("Project_ID") %>, '<% =Server.URLEncode(rs2("Project_Name")) %>')"  onMouseOver="top.status='Edit project Sarks.'; return true;" onMouseOut="top.status=''; return true;">Sark</a>]
<% else %>
			[Sark]
<% end if %>
			:</font></td>
		<td width=30% valign=top><font size=1 face="ms sans serif, arial, geneva" color=black>
<%
set rs3 = DBQuery("SELECT Employee.EmployeeID, Employee.FirstName, Employee.LastName, Employee.Username, Employee_Project_Xref.Employee_Desc FROM Employee_Project_Xref INNER JOIN Employee ON Employee_Project_Xref.Employee_ID = Employee.EmployeeID WHERE (Employee_Project_Xref.Project_ID = " & rs2("Project_ID") & ") and Employee.DateTermination is null and Employee.Branch_ID = " & Application("DefaultBranchID") & " ORDER BY Employee.LastName")
while not rs3.eof
	response.write("<a href='../../directory/content/details.asp?EmpID=" & rs3("EmployeeID") & "'>" & rs3("FirstName") & " " & rs3("LastName") & "</a>")
	if not isnull(rs3("Employee_Desc")) then response.write(" - " & rs3("Employee_Desc"))
	response.write("<br>")
	rs3.movenext
	wend
rs3.close
%>
			</font></td>
		<td width=10% align=right valign=top><font size=1 face="ms sans serif, arial, geneva" color=black>
<% if AddProjects then%>
			[<a href="javascript: AddTechs(<% =request("ClientID") %>, <% =rs2("Project_ID") %>, '<% =Server.URLEncode(rs2("Project_Name")) %>')"  onMouseOver="top.status='Edit project technologies.'; return true;" onMouseOut="top.status=''; return true;">Tech</a>]
<% else %>
			[Tech]
<% end if %>
			:</font></td>
		<td width=50% valign=top colspan=2 nowrap><font size=1 face="ms sans serif, arial, geneva" color=black>
<%
set rs3 = DBQuery("SELECT Tech.Tech_ID, Tech_Area.Tech_Area, Tech.Tech_Name FROM Project_Tech_Xref INNER JOIN Tech ON Project_Tech_Xref.Tech_ID = Tech.Tech_ID INNER JOIN Tech_Area ON Tech.Tech_Area_ID = Tech_Area.Tech_Area_ID WHERE (Project_Tech_Xref.Project_ID = " & rs2("Project_ID") & ") ORDER BY Tech_Area.Tech_Area, Tech.Tech_Name")
while not rs3.eof
	response.write("<a href='../../technology/content/tech.asp?techid=" & rs3("Tech_ID") & "&'>" & rs3("Tech_Area") & ", " & rs3("Tech_Name") & "</a><br>")
	rs3.movenext
	wend
rs3.close
if not isnull(rs2("OtherTech")) then response.write(Server.HTMLEncode(rs2("OtherTech")))
%>
			</font></td>
		</tr>

<%
		rs2.movenext
		if not rs2.eof then response.write("	<tr><td colspan=5 bgcolor=gray height=></td></tr>")
		wend
	end if
response.write("</table><p>")
rs.movenext
wend

rs.close
set rs = nothing
rs2.close
set rs2 = nothing
set rs3 = nothing
%>


<!-- #include file="../../footer.asp" -->
