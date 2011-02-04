<!-- #include file="../script.asp" -->
<%
'----------------------
'  SECTION PROPERTIES  
'----------------------
sectionTitle = "Technology"
sectionDir = "technology"

'----------------------
'  BUILD NAVIGATION    
'----------------------
buildNavItem "About this section", "about.asp"
buildNavItem "Knowledge Summary", "Knowledge_snapshot.asp"
buildNavItem "User Groups", "usergroups.asp"
sql = "SELECT Tech_Area.Tech_Area, Tech.Tech_Name, Tech.Tech_ID " & _
		"FROM Tech INNER JOIN Tech_Area ON Tech.Tech_Area_ID = Tech_Area.Tech_Area_ID " & _
		"INNER JOIN Tech_Area_Section ON Tech_Area.Tech_Area_Section_ID = Tech_Area_Section.Tech_Area_Section_ID " & _
		"WHERE Tech.Approved = 1" & _
		"ORDER BY Tech.Tech_Name, Tech_Area.Tech_Area"
set rs = DBQuery(sql)
while not rs.eof    navItem = ""
    if not IsNull(rs("tech_name")) then
	 navItem = rs("tech_name")
	 if not IsNull(rs("tech_area")) then	  navItem = navItem & " (" & rs("tech_area") & ")"
	 end if
	else
	 navItem = rs("tech_area")    end if
	buildNavItem navItem, "tech.asp?techid=" & rs("tech_id") & "&"
	rs.movenext
	wend
rs.close
%>
<!-- #include file="../header.asp" -->
