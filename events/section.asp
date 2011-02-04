<!-- #include file="../script.asp" -->
<%
'----------------------
'  SECTION PROPERTIES  
'----------------------
sectionTitle = "Events"
sectionDir = "events"
'------------------------------------------------
'  Strip off specific date when setting default  
'------------------------------------------------
x = instr(DefaultNavPage, "?date=")
if x > 0 then DefaultNavPage = left(DefaultNavPage, x+5)
'--------------------------
'  Build navigation items  
'--------------------------
buildNavItem "About this section", "about.asp"
buildNavItem "Links Around Town", "aroundtown.asp"
buildNavItem "Calendar", "calendar.asp?date=" & request.querystring("date")
buildNavItem "Events by Date", "bydate.asp?date=" & request.querystring("date")
buildNavItem "Events by Type", "bytype.asp?date=" & request.querystring("date")

if not Session("isGuest") then
 buildNavItem "My Events", "myevents.asp"
 buildNavItem "Add Events", "addnew.asp"
end if
%>
<!-- #include file="../header.asp" -->
