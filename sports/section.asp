<!-- #include file="../script.asp" -->
<!-- #include file="./content/sports_utils.asp" -->
<%
'----------------------
'  SECTION PROPERTIES  
'----------------------
sectionTitle = "Sports"
sectionDir = "sports"


buildNavItem "About", "default.asp"
if hasRole("WebMaster") then
 buildNavItem "Add/Remove Sport", "addSport.asp"
 buildNavItem "Assign Managers", "assignManager.asp"
end if
if isManager() then 
 buildNavItem "Sport Maintenance", "editSport.asp"
end if

sql="SELECT * FROM Sports ORDER BY Sport_Name"
set rsSports = DBQuery(sql)
do while not rsSports.eof
 buildNavItem rsSports("Sport_Name"), "showSport.asp?sportID="& rsSports("Sport_ID")
 rsSports.moveNext
loop

buildNavItem "Racquetball", "../../racquetball/content/about.asp"
buildNavItem "Wellness", "../../wellness/content/default.asp"

%>
<!-- #include file="../header.asp" -->
