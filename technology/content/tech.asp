<!--
Developer:    KDILL
Date:         09/01/1998
Description:  Technology template for current technologies.
-->


<!-- #include file="../section.asp" -->


<%
dim techid, valid, fIsExpert, Recipients, expertList, techLeadList, desc, tech_area

techid = request("techid")
valid = (techid <> "")
fIsExpert = False			'Set expert flag to default of false

SetWebMaster "cdolan"
'SectionBtns = SectionBtns & "<p><center><form><input type=button class=button value='Question...' onClick='SendEmail()'onMouseOver=" &chr(34)& " statmsg('Ask the experts a questions.'); return true;" & chr(34) & "onMouseOut=" & chr(34) & "statmsg('')"& chr(34) & "></form></center>"

if valid then
	'-----------------------------
	'  QUERY ON SUBJECT EXPERTS   
	'-----------------------------
	Recipients = ""
	set rs = DBQuery("select e.EmployeeID, e.FirstName, e.LastName, e.Username from employee e, employee_tech_xref x where x.Employee_ID = e.EmployeeID and x.Tech_ID = " & techid & " and DateTermination is null ORDER BY e.LastName, e.FirstName")
	if not rs.eof then
		while not rs.eof
			'Check if current user is expert and set flag
			'If UCase(Session("Username"))= UCase(rs("Username")) Then fIsExpert = True
			'Add Expert to list
			expertList = expertList & ", <a href='../../directory/content/details.asp?EmpID=" & rs("EmployeeID") & "'>" & rs("FirstName") & "&nbsp;" & rs("LastName") & "</a>"
			Recipients = Recipients & "," & lcase(rs("Username"))
			rs.movenext
			wend
		expertList = (mid(expertList,3))
		Recipients = mid(Recipients,2)
		rs.close
		end if
	'-----------------------------
	'  QUERY ON TECH DESCRIPTION  
	'-----------------------------
	set rs = DBQuery("select tech_desc, tech_area_id from tech where tech_id=" & techid)
	valid = not rs.eof
	end if

if valid then
	'------------------------------
	'  WRITE OUT TECH DESCRIPTION  
	'------------------------------
	desc = rs("tech_desc")
	desc = replace(desc, "&lt;", "<")
	desc = replace(desc, "&gt;", ">")
	if isnull(desc) or desc="" then desc = "[Description not available]"
	tech_area = rs("tech_area_id")
	rs.close
	'-----------------------------
	'  QUERY ON TECH LEADS        
	'-----------------------------

'SELECT      Tech.*, Tech_Area.Tech_Area, Tech_Area_Section.Section_Name, Employee.FirstName,
'                   Employee.LastName, Employee.Username
'FROM         Employee INNER JOIN Tech_Specialists ON Employee.EmployeeID 
'                       = Tech_Specialists.Employee_ID RIGHT OUTER JOIN Tech INNER JOIN Tech_Area ON 
'                       Tech.Tech_Area_ID = Tech_Area.Tech_Area_ID INNER JOIN Tech_Area_Section ON 
'                       Tech_Area.Tech_Area_Section_ID = Tech_Area_Section.Tech_Area_Section_ID ON 
'                       Tech_Specialists.Tech_Area_ID = Tech_Area.Tech_Area_ID
'WHERE      (Tech_Specialists.Tech_Specialist_Type_ID = 1)

	sql = "SELECT      Tech_Area.Tech_Area, Tech_Area_Section.Section_Name, Employee.EmployeeID, Employee.FirstName, " & _
			"                   Employee.LastName, Employee.Username, Tech_Specialist_Type.Specialist_Type " & _
			"FROM         Employee INNER JOIN Tech_Specialists ON Employee.EmployeeID " & _
			"                       = Tech_Specialists.Employee_ID INNER JOIN Tech_Specialist_Type ON " & _
			"                       Tech_Specialists.Tech_Specialist_Type_ID " & _
			"                       = Tech_Specialist_Type.Tech_Specialist_Type_ID INNER JOIN Tech_Area INNER JOIN " & _
			"                       Tech_Area_Section ON Tech_Area.Tech_Area_Section_ID " & _
			"                       = Tech_Area_Section.Tech_Area_Section_ID ON Tech_Specialists.Tech_Area_ID " & _
			"                       = Tech_Area.Tech_Area_ID " & _
			"WHERE      (Tech_Area.Tech_Area_ID = " & tech_area & ") " & _
			"ORDER BY Tech_Specialist_Type.Tech_Specialist_Type_ID, Employee.LastName, Employee.FirstName"
	set rs = DBQuery(sql)
	OldSpecialistType = ""
	while not rs.eof
		if rs("specialist_type") <> OldSpecialistType and OldSpecialistType <> "" then techLeadList = techLeadList & "<br>"
		if rs("specialist_type") <> OldSpecialistType then
			techLeadList = techLeadList & "<font color=navy><b>" & rs("tech_area") & " " & rs("specialist_type") & "</b></font>:&nbsp;&nbsp;"
		else
			techLeadList = techLeadList & ", "
		end if
		techLeadList = techLeadList & "<a href='../../directory/content/details.asp?EmpID=" & rs("EmployeeID") & "'>"
		techLeadList = techLeadList & rs("FirstName") & " " & rs("LastName") & "</a>"
		techLeadList = techLeadList & NL & NL
		OldSpecialistType = rs("specialist_type")
		If UCase(Session("Username"))= UCase(rs("Username")) OR hasRole("WebMaster") Then fIsExpert = True
		rs.movenext
		wend
	rs.close
	techLeadList = techLeadList & "<p>"
%>

<%=NL & NL & techLeadList%>

<font color=navy><b>Subject matter experts:</b></font>&nbsp;&nbsp;<%=expertList%><p>

<table width="100%" border=0 cellpadding=0 cellspacing=0><tr>
<%Response.Write("<td width='99%'><font size=1 face='ms sans serif, arial, geneva'>" & desc & "</font></td>")%>
	<td width="1%" nowrap align=right valign=bottom>
<%if fIsExpert then%>
		<font size=1 face="ms sans serif, arial, geneva">
			[<a href="javascript:EditDesc();" onMouseOver="top.status='Edit Description.'; return true;" onMouseOut="top.status=''; return true;">Edit</a>]
		</FONT>
<%end if%>
	</td>
</tr></table><hr size=1><br>


<table width="100%" border=0 cellpadding=0 cellspacing=0>
<tr>
	
	<td width="53%" valign=top>
	<!-- #include file="link.asp" -->
	</td>

	<td><img src="../../common/images/spacer.gif" height=1 width=20 border=0></td>
	
	<td width="47%" valign=top>
	<!-- #include file="tech_projects.asp" -->
	</td>

</tr></table>

<%
	else
		response.write("Invalid techid!")
	end if
%>

<script language=javascript>
<!--
function EditDesc(){winTech=window.open("editdesc.asp?empname=<%=Session("Username")%>&Tech_ID=<%=techid%>", "EditDesc","height=300,width=600,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1, scrollbar=1")}
function SendEmail(){winTech=window.open("../../email.asp?recipient=<%=Server.UrlEncode(Recipients)%>&subject=<%=Server.URLEncode("QUESTION: " & pageTitle)%>", "SendEmail","height=300,width=520,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1, scrollbar=1")}
// -->
</script>


<!-- #include file="../../footer.asp" -->
