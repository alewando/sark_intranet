<!-- #include file="../script.asp" -->
<%
'----------------------
'  SECTION PROPERTIES  
'----------------------
sectionTitle = "Racquetball League"
sectionDir = "racquetball"

racquetballStartDate = "11/06/2000"
racquetballEndDate   = "12/31/2000"

if hasRole("racquetballAdmin") then 
 buildNavItem "Edit Leagues", "editSchedule.asp"
end if
buildNavItem "About", "about.asp"
buildNavItem "Rules & Regulations", "rules.asp"
buildNavItem "Schedule", "schedule.asp"


sql="SELECT * FROM Racquetball_League ORDER BY League_ID"
set rs=DBQuery(sql)
do while not rs.eof
 buildNavItem rs("Description"), "showSchedule.asp?leagueID=" & rs("League_ID")
 rs.movenext
loop
rs.close
 
%>
<!-- #include file="../header.asp" -->
