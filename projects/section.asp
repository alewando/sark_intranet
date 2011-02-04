<!-- #include file="../script.asp" -->
<%
'----------------------
'  SECTION PROPERTIES  
'----------------------
if sectionTitle = "" then sectionTitle = "Projects"
sectionDir = "projects"

buildNavItem "About this section", "about.asp"
buildNavItem "Client Summaries", "clients.asp"


if (Session("ID")) then
	set rs = DBQuery("SELECT Client_ID FROM Employee  WHERE EmployeeID = " & Session("ID"))
		buildNavItem "Current Client", "details.asp?clientid=" & rs("Client_ID") 
end if




if session("OfficeStaff") then buildNavItem "Client / Project Report!", "details.asp"

sql = "SELECT Client_ID, ClientName FROM Client WHERE SortOverride=0 AND Client_ID > 2 ORDER BY ClientName"
set rs = DBQuery(sql)
while not rs.eof
	buildNavItem rs("ClientName"), "details.asp?clientid=" & rs("Client_ID") & "&"
	rs.movenext
	wend
rs.close
%>

<!-- #include file="../header.asp" -->
