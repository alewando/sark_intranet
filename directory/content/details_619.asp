<!--
Developer:		GINGRICE
Date:			08/18/1998
Description:	Displays detailed listing of consultants
History:		12/11/1998 - KDILL - Added support for email display.-->


<% pageTitle = "Details" %>

<!-- #include file="../section.asp" -->

<SCRIPT language = javascript>
<!--function OpenWin(page, EmpID, WinTitle){
	winTech= window.open(page + "?empID=" + EmpID, WinTitle,"height=615,width=500,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1,scrollbars=1")
	}

// --></SCRIPT>

<%
'-------------------------------------
'  Get Reference to FileSystemObject  
'-------------------------------------
Set FileObject = Server.CreateObject("Scripting.FileSystemObject")

'---------------------------------------------
'	Execute Database Query for Employee Info  
'---------------------------------------------
ls_sql = "select e.*, t.title_desc  from employee e, employee_title t where e.employeetitle_id *= t.employee_title_id and e.employeeid = " & Request.QueryString("EmpID") 
set rs = DBQuery(ls_sql)
curTitle = curTitle & " (" & rs("firstname") & " " & rs("lastname") & ")"
if IsNull(rs("title_desc")) then
	ls_title = ""
else
	ls_title = rs("title_desc")
end if

'--------------------------
'	Display Employee Name  
'--------------------------
Response.Write("<center><font color=navy><b>" & rs("firstname") & " " & rs("lastname") & ",&nbsp;&nbsp;" & ls_title & "</b></font></center><br>")
Response.Write("<TABLE BORDER=0 WIDTH='100%'>")
ls_clientphone = rs("client_phone")
ldt_startdate = rs("startdate")
ls_username = lcase(rs("username"))
ls_empid = rs("employeeid")
ls_pager = rs("pagernumber")ls_pager_email = trim(rs("pageremail"))
ls_cellphone = rs("cellphone")
ls_voicemail = rs("voicemail")ls_fax = rs("fax")

'-----------------------------------
'	Check if Employee Picture Exists
'-----------------------------------
imgEmployee = "images/" & rs("username") & ".jpg" 
FileName = Server.MapPath (imgEmployee)
if NOT FileObject.FileExists(FileName) then imgEmployee = "images/no_photo.jpg"
Response.Write("<TR><TD ALIGN=Center VALIGN=Top ROWSPAN=2><img src='" & imgEmployee & "' height=100 width=100></TD>")

'-------------------------
'	Display Personal Info
'-------------------------
Response.Write("<TD VALIGN=TOP><FONT SIZE=1 face='ms sans serif, arial, geneva'><b>" & "Home Address:" & "</b><br>")
if NOT rs("unlistedpersonal") then
	Response.Write(rs("address") & "<br>" & rs("city") & "," & rs("state") & "  " & rs("zip") & "<br>" & rs("homephone") & "<br><br></TD>")
else
	Response.Write("Unlisted<br><br></TD>")
end if
Response.Write("<TD VALIGN=TOP><FONT SIZE=1 face='ms sans serif, arial, geneva'><b>" & "Personal Information:" & "</b><br>")
if rs("spousename") <> "" then
	Response.Write("Spouse: ")
	Response.Write(rs("spousename")) & "<br>" 
end if
Response.Write("Birthday: ")
if not isnull(rs("birthday")) then
	if isdate(rs("birthday")) then
		Response.Write(MonthName(DatePart("m", rs("birthday"))) & " " & DatePart("d", rs("birthday"))) & "<br>"
	end if
end if
if rs("PersonalWebsite") <> "" then
	Response.Write("Website: ")
	Response.Write("<a href='http://" & rs("PersonalWebsite") & "'> " & rs("PersonalWebsite") & "</a>")
end if
'---------------------------------------------
'	Execute Database Query for Client Info    
'---------------------------------------------
if not isnull(rs("client_id")) then
	ls_sql = "select * from client where client_id= " & rs("client_id")
	set rs = DBQuery(ls_sql)
	ls_clientname = rs("clientname")
	ln_sortoverride = rs("sortoverride")
	end if
Response.Write("</TD></TR><TR><TD VALIGN=TOP><FONT SIZE=1 face='ms sans serif, arial, geneva'>")

'-------------------------
'	Display SARK Info     
'-------------------------
Response.Write("<b>" & "SARK Information: " & "</b><br>")
if ls_username <> "" then Response.Write("Email: <a href='javascript:SendEmail();'>" & ls_username & "@sark.com</a><br>")
Response.Write("Start Date: " & ldt_startdate & "<br><br></td>")
if ln_sortoverride <> 2 then
	Response.Write("<TD VALIGN=TOP><FONT SIZE=1 face='ms sans serif, arial, geneva'><b>" & "Client:" & "</b><br>")
	Response.Write(ls_clientname & "<br>" & ls_clientphone)
	Response.Write("</TD></TR>")
end if

'if NOT IsNull(ls_voicemail) then
	' Display office phone numbers
	Response.Write("<TR><TD></TD><TD VALIGN=TOP><FONT SIZE=1 face='ms sans serif, arial, geneva'>")
	Response.Write("<b>" & "Other Numbers: " & "</b><br>")
	Response.Write("Pager: " & ls_pager & "<br>")	if ls_pager_email <> "" then
		Response.Write("Pager Email: <a href='javascript:SendPage();'>" & ls_pager_email & "</a><br>")	end if
	Response.Write("Mobile: " & ls_cellphone & "<br>")
	Response.Write("Voice Mail: " & ls_voicemail & "<br>")	Response.Write("Fax: " & ls_fax & "<br>")
	Response.Write("</TD></TR></TABLE>")
'end if
'**********
Response.Write(vbNewLine & vbNewLine & "<!--" & vbNewLine & _               "You're looking at   " & trim(ls_username) & vbNewLine & _
               "You're logged in as " & Session("Username") & vbNewline & _
               "-->" & vbNewLine & vbNewLine)'**********
'---------------------------------------
'Add Button to update employee info
'---------------------------------------
ls_empid = Request.QueryString("EmpID")
if ( UCase(trim(ls_username)) = UCase(session("Username")) ) or _ 
		 (UCase(Application("WebMaster")) = UCase(session("Username"))) then	Response.Write("<TABLE BORDER=0 WIDTH='100%'><TR><form>")
	Response.Write("<TD WIDTH='50%'><INPUT TYPE=button class=button NAME='UpdateEmp' VALUE = 'Edit Employee' onClick=" & chr(34) & "OpenWin('editemp.asp', " & ls_empid  & ", 'EditSARKEmp')" & chr(34) & "></TD>")
	Response.Write("</form></TR>")	Response.Write("</TABLE>")
end if

'--------------------------------------------------
'	If details page belongs to current user, then 
'	display edit button
'--------------------------------------------------
ls_empid = Request.QueryString("EmpID")Response.Write("<TABLE BORDER=0 WIDTH='100%'>")Response.Write("<TR><TD COLSPAN=4><hr size=1></TD></TR></table>")

Response.Write("<TABLE BORDER=0 WIDTH='75%'>")
'---------------------------------------------------------------------------------
' If details page belongs to current user then display edit buttons.
'---------------------------------------------------------------------------------if ( UCase(trim(ls_username)) = UCase(session("Username")) ) or _
	( UCase(Application("WebMaster")) = UCase(session("Username")) ) then	Response.Write("<TR HEIGHT='35' ALIGN='left' VALIGN='TOP'><form>")
	if (UCase(trim(ls_username)) = UCase(session("Username"))) then				'----------------------------------------------------------------		' Edit SARK Profile                                    
		'----------------------------------------------------------------
		
		Response.Write("<TD><INPUT TYPE=button class=button NAME='UpdateProfile' " & _					   "VALUE = '   Edit Profile  ' onClick=" & chr(34) & _					   "OpenWin('editprofile.asp', " & ls_empid  & ", 'EditSARKProfile')" & chr(34) & "></TD>")		
		'----------------------------------------------------------------		' Edit Technical Expertise                                    
		'----------------------------------------------------------------
				Response.Write("<TD><INPUT TYPE=button class=button NAME='TechExpert' " & _		               "VALUE = ' Edit Expertise ' onClick=" & chr(34) & _		               "OpenWin('techexperts.asp', " & ls_empid  & ", 'EditSARKTech')" & chr(34) & "></TD>")		'----------------------------------------------------------------		' Edit Industry Experience                                    
		'----------------------------------------------------------------

		Response.Write("<TD><INPUT TYPE=button class=button NAME='IndExperience' " & _		               "VALUE = '  Edit Industry  ' onClick=" & chr(34) & _		               "OpenWin('indexperience.asp', " & ls_empid  & ", 'EditSARKInd')" & chr(34) & "></TD>")
		'----------------------------------------------------------------		' Edit Skills Inventory                                    
		'----------------------------------------------------------------

		Response.Write("<TD><INPUT TYPE=button class=button NAME='SkillsInv' " & _		               "VALUE = '  Edit Skills Inventory  ' onClick=" & chr(34) & _		               "OpenWin('SkillsInv.asp', " & ls_empid  & ", 'EditSARKInd')" & chr(34) & "></TD>")
	end if

	'----------------------------------------------------------------	' Edit Certifications                                    
	'----------------------------------------------------------------
	if ( UCase(Application("WebMaster")) = UCase(session("Username")) ) then		Response.Write("<TD WIDTH='37%'><INPUT TYPE=button class=button NAME='Certifictaions' " & _					   "VALUE = 'Edit Certifications' onClick=" & chr(34) & _		               "OpenWin('certifications.asp', " & ls_empid  & ", 'EditSARKCert')" & chr(34) & "></TD>")		               		if ( UCase(trim(ls_username)) <> UCase(session("Username")) ) then
			Response.Write("<TD><INPUT TYPE=button class=button NAME='IndExperience' " & _			               "VALUE = '  Edit Industry  ' onClick=" & chr(34) & _			               "OpenWin('indexperience.asp', " & ls_empid  & ", 'EditSARKInd')" & chr(34) & "></TD>")	
		end if	end if
	
	Response.Write("</form></TR>")
end if

Response.Write("</table>")
Response.Write("<TABLE BORDER=0 WIDTH='100%'>")
'--------------------------------------------
'	Execute Database Query for Profile Info  
'--------------------------------------------ls_sql = "select * from profile where employee_id = " & ls_empid
set rs = DBQuery(ls_sql)
    Response.Write("<TR><TD WIDTH='60%' ALIGN=LEFT VALIGN=TOP rowspan=5>")
 	Response.Write("<table width='100%' cellspacing=0 cellpadding=4 border=0 bgcolor=#ffffcc class=tableShadow>")	Response.Write("<tr><td><font size=1 face=ms sans serif, arial, geneva color=navy><b>Profile</b></font></td></tr>") 	
 	Response.Write("<tr><td colspan=2 bgcolor=gray height=1></td></tr>")
 		if  rs.eof then
	'--------------------------------------------------
	'	Display Profile Info - Only display is not null
	'--------------------------------------------------
	 	Response.Write("<tr><td><font size = 1 color = gray>None Specified</font></td></tr>")
 	
else
	
	Response.Write("<tr><td>")
	Response.write("<font size=1 face='ms sans serif, arial, geneva'>")
	if not (isnull(rs("college")) or trim(rs("college")) = "") then
		Response.Write("<b>" & "College:" & "<br></b>" & rs("college") & "<br><br>")
		end if
	if not (isnull(rs("hobbies_sports_interests")) or trim(rs("hobbies_sports_interests")) = "" ) then
		Response.Write("<b>" & "Hobbies, Sports & Interests:" & "<br></b>" & rs("hobbies_sports_interests") & "<br><br>")
		end if
	if not (isnull(rs("dream_vacation")) or trim(rs("dream_vacation")) = "" ) then
		Response.Write("<b>" & "Dream Vacation:" & "<br></b>" & rs("dream_vacation") & "<br><br>")
		end if
	if not (isnull(rs("favorite_show_ad")) or trim(rs("favorite_show_ad")) = "" ) then
		Response.Write("<b>" & "Favorite Show/Ad:" & "<br></b>" & rs("favorite_show_ad") & "<br><br>")
		end if
	if not (isnull(rs("person_inspired")) or trim(rs("person_inspired")) = "" ) then
		Response.Write("<b>" & "Person Who Inspired me the Most:" & "<br></b>" & rs("person_inspired") & "<br><br>")
		end if
	if not (isnull(rs("favorite_motto")) or trim(rs("favorite_motto")) = "" ) then
		Response.Write("<b>" & "Favorite Motto:" & "<br></b>" & rs("favorite_motto") & "<br><br>")
		end if
	if not (isnull(rs("funniest_embarrassing")) or trim(rs("funniest_embarrassing")) = "" ) then
		Response.Write("<b>" & "Funniest or Most Embarrassing Moment:" & "<br></b>" & rs("funniest_embarrassing") & "<br><br>")
		end if 
	if not (isnull(rs("lotto")) or trim(rs("lotto")) = "" ) then
		Response.Write("<b>" & "If I Won the Lotto...:" & "<br></b>" & rs("lotto") & "<br><br>")
		end if
	if not (isnull(rs("pet_peeve")) or trim(rs("pet_peeve")) = "" ) then
		Response.Write("<b>" & "Pet Peeve:" & "<br></b>" &  rs("pet_peeve") & "<br><br>")
		end if
	if not (isnull(rs("prized_possession")) or trim(rs("prized_possession")) = "" ) then
		Response.Write("<b>" & "Prized Possession:" & "<br></b>" & rs("prized_possession") & "<br><br>")
		end if
	if not (isnull(rs("biggest_challenge")) or trim(rs("biggest_challenge")) = "" ) then
		Response.Write("<b>" & "Biggest Challenge:" & "<br></b>" & rs("biggest_challenge") & "<br><br>")
		end if
	Response.write("</font>")
	end if
	Response.Write("</td></tr></table>")rs.close

Response.write("<TD>&nbsp;</TD>")
Response.Write("<td ALIGN=LEFT VALIGN=top><font size=1 face='ms sans serif, arial, geneva'>")

	Response.Write("<table width='100%' cellspacing=0 cellpadding=4 border=0 bgcolor=#ffffcc class=tableShadow>")	Response.Write("<tr><td><font size=1 face=ms sans serif, arial, geneva color=navy><b>Technologies</b></font></td></tr>")
	Response.Write("<tr><td colspan=2 bgcolor=gray height=1></td></tr>")
	Response.Write("<tr><td><font size=1 face= ms sans serif, arial, geneva color=black>")

sql = "SELECT Tech_Area.Tech_Area, Tech.Tech_Name, Employee_Tech_Xref.Employee_ID, Tech.Tech_ID" _
		& " FROM Tech INNER JOIN Employee_Tech_Xref ON Tech.Tech_ID = Employee_Tech_Xref.Tech_ID " _
		& " INNER JOIN Tech_Area ON Tech.Tech_Area_ID = Tech_Area.Tech_Area_ID " _
		& " WHERE (Employee_Tech_Xref.Employee_ID = " & ls_empid & ") AND" _
		& " (Tech.Approved = 1)" _
		& " ORDER BY Tech_Area.Tech_Area, Tech.Tech_Name"
set rs = DBQuery(sql)
	if rs.eof then	
	Response.Write("<tr><td><font size = 1 color = gray>None Specified</font></td></tr>")else        Dim HoldTech     
    
   	HoldTech = rs("Tech_Area")       Response.Write("<b>" & HoldTech & "</b><br>")
        while not rs.eof
		if rs("Tech_Area") <> HoldTech then
			HoldTech = rs("Tech_Area")			Response.Write("<b>" & HoldTech & "</b><br>")
		end if													
		response.write("&nbsp;&nbsp;<A HREF='../../technology/content/tech.asp?techid=" & rs("Tech_ID") & "'>" & rs("Tech_Name") & "</a><br>")		rs.movenext

	wendend if	

Response.Write("</font></td></tr></table>")
Response.Write("<table><tr><td></td></tr></table>")
Response.Write("<table width=100% cellspacing=0 cellpadding=4 border=0 bgcolor=#ffffcc class=tableShadow>")Response.Write("<tr><td><font size=1 face=ms sans serif, arial, geneva color=navy><b>Certifications</b></font></td></tr>")
Response.Write("<tr><td colspan=2 bgcolor=gray height=1></td></tr>")
Response.Write("<tr><td><font size=1 face= ms sans serif, arial, geneva color=black>")

'SQL statement for gathering industry experience information.
sql = "SELECT Certifications.Cert_Name" _
	& " FROM Employee_Cert_Xref INNER JOIN" _
	& " Certifications ON" _
    & " Employee_Cert_Xref.Cert_ID = Certifications.Cert_ID" _
	& " WHERE (Employee_Cert_Xref.Employee_ID = " & ls_empid & ")" _
	& " ORDER BY Certifications.Cert_Name"			
set rs = DBQuery(sql)

if rs.eof then
	Response.Write("<tr><td><font size = 1 color = gray>None Specified</font></td></tr>")else    while not rs.eof
		Response.Write("<b>" & rs("Cert_Name") & "</b><br>")
		rs.movenext

	wendend if	

rs.close
Response.Write("</font></td></tr></table>")Response.Write("<table><tr><td></td></tr></table>")
'******************************************************************'*  Button to open a new window where the user can edit his/her   *'*  industry experience (03/09/2000 - DJM)                        *
'******************************************************************

Response.Write("<table width=100% cellspacing=0 cellpadding=4 border=0 bgcolor=#ffffcc class=tableShadow>")Response.Write("<tr><td><font size=1 face=ms sans serif, arial, geneva color=navy><b>Industry Experience</b></font></td></tr>")
Response.Write("<tr><td colspan=2 bgcolor=gray height=1></td></tr>")
Response.Write("<tr><td><font size=1 face= ms sans serif, arial, geneva color=black>")

'SQL statement for gathering industry experience information.
sql = "SELECT Industry.Industry_Name" _
		& " FROM Employee_Industry_Xref INNER JOIN" _
		& " Industry ON" _
		& " Employee_Industry_Xref.Industry_ID = Industry.Industry_ID" _
		& " WHERE (Employee_Industry_Xref.Employee_ID = " & ls_empid & ")" _
		& " ORDER BY Industry.Industry_Name"			
set rs = DBQuery(sql)

if rs.eof then
	Response.Write("<tr><td><font size = 1 color = gray>None Specified</font></td></tr>")else    while not rs.eof
		Response.Write("<b>" & rs("Industry_Name") & "</b><br>")
		rs.movenext

	wendend if	

rs.close
Response.Write("</font></td></tr></table>")Response.Write("<table><tr><td></td></tr></table>")


Response.Write("<table width=100% cellspacing=0 cellpadding=4 border=0 bgcolor=#ffffcc class=tableShadow>")Response.Write("<tr><td><font size=1 face=ms sans serif, arial, geneva color=navy><b>Projects</b></font></td></tr>")
Response.Write("<tr><td colspan=2 bgcolor=gray height=1></td></tr>")
Response.Write("<tr><td><font size=1 face= ms sans serif, arial, geneva color=black>")

'SQL statement for gathering the projects worked on.

ProjectSql="SELECT Employee_Project_Xref.Employee_ID, " _
    & "Employee_Project_Xref.Project_ID, Client.Client_ID, " _
    & "Client.ClientName, Project.Project_Name " _
	& "FROM Client INNER JOIN " _
    & "Project ON Client.Client_ID = Project.Client_ID INNER JOIN " _
    & "Employee_Project_Xref ON " _
    & "Project.Project_ID = Employee_Project_Xref.Project_ID " _
	& "WHERE (Employee_Project_Xref.Employee_ID = " & ls_empid & ")"

set ProjectRs=DBQuery(ProjectSql)

if ProjectRs.eof then
	Response.Write("<tr><td><font size=1 color = gray>None Specified</font></td></tr>")
else
	while not ProjectRs.eof
		Response.Write("<a href='../../projects/content/details.asp?clientid=" & ProjectRs("Client_ID") & "'>" & ProjectRs("ClientName") & "</a> - <a href='../../projects/content/details.asp?clientid= " & ProjectRs("Client_ID") & "&projectid=" & ProjectRs("Project_ID") & "'>" & ProjectRs("Project_Name") & "</a><br>")
		ProjectRs.movenext
	wend
end if

ProjectRs.close
Response.Write("</font></td></tr></table>")Response.Write("</font></TD></TR></TABLE>")
%>

<script language=javascript>
<!--
function SendEmail(){winTech= window.open("../../email.asp?editto=yes&recipient=<%=ls_username%>&footer=<%=Server.URLEncode(UCase(Application("web")) & ": " & curTitle)%>", "SendEmail","height=320,width=520,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1, scrollbar=1")}
function SendPage(){winTech= window.open("../../pager.asp?recipient=<% =ls_pager_email %>&from=<%=session("Username")%>&name=<%=Server.URLEncode(session("name"))%>&footer=<%=Server.URLEncode(UCase(Application("web")) & ": " & curTitle)%>", "SendPage","height=320,width=520,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1, scrollbar=1")}
// -->
</script>

<!-- #include file="../../footer.asp" -->
