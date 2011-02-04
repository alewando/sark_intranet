<!-- #include file="../script.asp" -->

<%
'----------------------
'  SECTION PROPERTIES  
'----------------------
ls_user=Session("ID")
'sectionTitle = "Welcome " & Session("firstname")
if not Session("isGuest") then
sectionTitle = "Welcome " & "<a target=body href='../../directory/content/details.asp?EmpID=" & Session("ID") & "'>" & Session("firstname") & "</a>"
else
sectionTitle = "Welcome Guest"
end if
sectionDir = "welcome"


if (Month(Date)=4) and (Day(Date)=1) then
	sectionNav = "<font face='ms sans serif, arial, geneva' color=red><b>April Fool's Day!</b></font>"
else
	sectionNav = "<font face='ms sans serif, arial, geneva'>" & WeekDayName(Weekday(date)) & ",&nbsp;&nbsp;" & MonthName(Month(Date)) & " " & Day(Date) & ", " & Year(Date) & "</font>"
end if
sectionBtns = ""
%>

<!-- #include file="../header.asp" -->

