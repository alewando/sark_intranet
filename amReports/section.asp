<!-- #include file="../script.asp" -->

<%
'----------------------
'  SECTION PROPERTIES  
'----------------------
sectionTitle = "A/M Reports"
sectionDir = "amReports"

if (hasRole("AccountManager")) then
	buildNavItem "About", "default.asp"
	buildNavItem "Vacation-30 days-alpha", "vacation_30d_alpha.asp"
	buildNavItem "Vacation-30 days-by date", "vacation_30d_bydate.asp"
	buildNavItem "Vacation-12 months", "vacation_year.asp"
	buildNavItem "Vacation-Past 2 weeks", "vacation_2weeks.asp"
end if

if hasRole("SponsorAdmin") or hasRole("AccountManager") then
	buildNavItem "Reviews by Date", "review_date.asp"
	buildNavItem "Reviews by A/M", "review_amDate.asp"
	buildNavItem "Reviews by A/M- Next 90 Days", "review_Date90.asp"
	buildNavItem "Reviews Past Due", "review_past_due.asp"
	buildNavItem "Sponsor/Sponsoree", "review_sponsorDate.asp"
end if

if hasRole("AccountManager") then	
	buildNavItem "Exams by Consultants", "consexams.asp"
	buildNavItem "Consultants by Exams", "examlist.asp"
	buildNavItem "Exams 0-60d Old", "exam_0-60_day.asp"
	buildNavItem "Exams by Vendor", "exams_by_vendor.asp"
	buildNavItem "Certifications by Category", "cert_by_cat.asp"
	buildNavItem "MCP Listing", "mcp_report.asp"
	buildNavItem "SARK's by Classes", "classestaken.asp"
	buildNavItem "Classes by SARK's", "consclasses.asp"
	buildNavItem "Skills Inv by Date", "skillsInv_date.asp"
	buildNavItem "Skills Inv - Not Entered", "list_skillsInv.asp"
	buildNavItem "Skills Inv - Aging", "skillsInv_stale.asp"
	buildNavItem "Consultants by Title", "consTitles.asp"
	buildNavItem "Client List with Sales Rep and A/M", "client_sr_am.asp"
	buildNavItem "Client List by A/M", "client_am.asp"
	buildNavItem "No Techs Listed", "no_techs.asp"
	buildNavItem "Intranet Survey", "intranet_survey.asp"
end if

if hasRole("WebMaster") then 
	buildNavItem "Employee Security", "employeeSecurity.asp"
end if
%>

<!-- #include file="../header.asp" -->
