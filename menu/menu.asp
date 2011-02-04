/*
 * Sark Cincinnati intranet menu
 * 
 * To add a new section to the menu, use the following syntax:
 *  aux1=insFld(mnuRoot, gFld("Section Name", linkPrefix+"path_to_section/content/", "./menu/section_name.gif"));
 *  Images should be 27x27
 *
 * To add new menu items, use the following syntax within the correct section:
 * insDoc(aux1, gLnk(2,"Page Name", linkPrefix+"path_to/page.asp"));
 *
 * Note the use of the hasRole("RoleName") and allowed(navSection, item) functions.
 * hasRole("RoleName") returns true if the currently logged in user has been granted
 *  the security role in question
 * allowed(navSection, item) returns true if the currently logged in user has selected
 *  the item in question in their preferences. (see /preferences/content/default.asp)
 */

linkPrefix="/intranet/";
//topOffset=document.images("imgLogo").height + 10;
topOffset=110;	// For NS4, offset (in pixels) from top of page to start of menu (adjustment for branch logo)

//mnuRoot=gFld("", linkPrefix+"welcome/content/", "./menu/logo_root.gif");
mnuRoot=gFld("", linkPrefix+"welcome/content/", "./menu/blank.gif");
aux1=insFld(mnuRoot, gFld("Welcome", linkPrefix+"welcome/content/", "./menu/welcome.gif"));
<% if allowed(navWelcome, preferences) then %>
insDoc(aux1, gLnk(2,"Preferences",linkPrefix+"preferences/content/default.asp"));
<% end if %>

<% if hasRole("WebMaster") or hasRole("TrainingCoordinator") then %>
aux1=insFld(mnuRoot, gFld("Tools", linkPrefix+"tools/content/", "./menu/tools.gif"));
 <% if hasRole("WebMaster") then %>
  insDoc(aux1, gLnk(2,"Add Employee", linkPrefix+"tools/content/add_emp.asp"));
  insDoc(aux1, gLnk(2,"Edit Employee", linkPrefix+"tools/content/chg_emp.asp"));
  insDoc(aux1, gLnk(2,"Delete Employee", linkPrefix+"tools/content/del_emp.asp"));
  insDoc(aux1, gLnk(2,"Edit Emp Security", linkPrefix+"tools/content/edit_user_role.asp"));
  insDoc(aux1, gLnk(2,"Add Skill to S/I", linkPrefix+"tools/content/add_skill.asp"));
  insDoc(aux1, gLnk(2,"Add Client", linkPrefix+"tools/content/add_client.asp"));
  insDoc(aux1, gLnk(2,"Edit Client", linkPrefix+"tools/content/edit_client.asp"));
  insDoc(aux1, gLnk(2,"Approve Postings", linkPrefix+"tools/content/approve_postings.asp"));
  insDoc(aux1, gLnk(2,"Approve New Tech", linkPrefix+"tools/content/techapproval.asp"));
  insDoc(aux1, gLnk(2,"Edit Tech", linkPrefix+"tools/content/edit_technology.asp"));
  insDoc(aux1, gLnk(2,"Add Certification", linkPrefix+"tools/content/add_cert.asp"));
  insDoc(aux1, gLnk(2,"Edit Certification", linkPrefix+"tools/content/edit_cert.asp"));
  insDoc(aux1, gLnk(2,"Add Exam", linkPrefix+"tools/content/add_exam.asp"));
  insDoc(aux1, gLnk(2,"Edit Exam", linkPrefix+"tools/content/edit_exam.asp"));
  insDoc(aux1, gLnk(2,"Add Course", linkPrefix+"tools/content/add_course.asp"));
 <% end if %>
 <% if hasRole("TrainingCoordinator") then %>
  insDoc(aux1, gLnk(2,"Add Class", linkPrefix+"tools/content/add_class.asp"));
  insDoc(aux1, gLnk(2,"Edit Class", linkPrefix+"tools/content/edit_class.asp"));
 <% end if %>
 <% if hasRole("WebMaster") then %>
  insDoc(aux1, gLnk(2,"Add Solution Services", linkPrefix+"tools/content/add_ss.asp"));
  insDoc(aux1, gLnk(2,"Edit Solution Services", linkPrefix+"tools/content/edit_ss.asp"));
  insDoc(aux1, gLnk(2,"Add Branch", linkPrefix+"tools/content/add_branch.asp"));
  insDoc(aux1, gLnk(2,"Edit Branch", linkPrefix+"tools/content/edit_branch.asp"));
  insDoc(aux1, gLnk(2,"Edit Announcements", linkPrefix+"tools/content/updateannouncement.asp?Announcement_ID=1"));
 <% end if %>
<% end if %>

<% if hasRole("AccountManager") or hasRole("SponsorAdmin") or hasRole("Sales") then %>
aux1 = insFld(mnuRoot, gFld("AM Reports", linkPrefix+"amReports/content/", "./menu/amreports.gif"));
 <% if hasRole("AccountManager") then %>
  aux2=insFld(aux1, gFld("Vacation", linkPrefix+"amReports/content/"));
  insDoc(aux2, gLnk(2, "30d out by date", linkPrefix+"amReports/content/vacation_30d_bydate.asp"));
  insDoc(aux2, gLnk(2, "30d out by alpha", linkPrefix+"amReports/content/vacation_30d_alpha.asp"));
  insDoc(aux2, gLnk(2, "1yr out alpha", linkPrefix+"amReports/content/vacation_year.asp"));
  insDoc(aux2, gLnk(2, "Past 2 weeks", linkPrefix+"amReports/content/vacation_2weeks.asp"));
 <% end if%>
 <% if hasRole("SponsorAdmin") or hasRole("AccountManager") then %>
  aux2=insFld(aux1, gFld("Reviews", linkPrefix+"amReports/content/"));
  insDoc(aux2, gLnk(2, "Rvws By Date", linkPrefix+"amReports/content/review_date.asp"));
  insDoc(aux2, gLnk(2, "Rvws By AM", linkPrefix+"amReports/content/review_amDate.asp"));
  insDoc(aux2, gLnk(2, "Rvws 90d out", linkPrefix+"amReports/content/review_date90.asp"));
  insDoc(aux2, gLnk(2, "Rvws Past Due", linkPrefix+"amReports/content/review_past_due.asp"));
  insDoc(aux2, gLnk(2, "Sponsor List", linkPrefix+"amReports/content/review_sponsorDate.asp"));
 <% end if %>
 <% if hasRole("Sales") or hasRole("AccountManager") then %>
  aux2=insFld(aux1, gFld("Exams &amp; Classes", linkPrefix+"amReports/content/"));
  insDoc(aux2, gLnk(2, "Exams by SARKs", linkPrefix+"amReports/content/consexams.asp"));
  insDoc(aux2, gLnk(2, "SARKs by Exams", linkPrefix+"amReports/content/examlist.asp"));
  insDoc(aux2, gLnk(2, "Exams 0-60d old", linkPrefix+"amReports/content/exam_0-60_day.asp"));
  insDoc(aux2, gLnk(2, "Exams by vendor", linkPrefix+"amReports/content/exams_by_vendor.asp"));
  insDoc(aux2, gLnk(2, "Certs by Cat.", linkPrefix+"amReports/content/cert_by_cat.asp"));
  insDoc(aux2, gLnk(2, "MCP Listing", linkPrefix+"amReports/content/mcp_report.asp"));
  insDoc(aux2, gLnk(2, "Classes by Sarks", linkPrefix+"amReports/content/consclasses.asp"));
  insDoc(aux2, gLnk(2, "Sarks by Classes", linkPrefix+"amReports/content/classestaken.asp"));
 <% end if %>
 <% if hasRole("AccountManager") then %>
  aux2=insFld(aux1, gFld("Skills Inventory", linkPrefix+"amReports/content/"));
  insDoc(aux2, gLnk(2, "Skills Inv: By Date", linkPrefix+"amReports/content/skillsInv_date.asp"));
  insDoc(aux2, gLnk(2, "Skills Inv: Tardy", linkPrefix+"amReports/content/list_skillsInv.asp"));
  insDoc(aux2, gLnk(2, "Skills Inv: 90d old", linkPrefix+"amReports/content/skillsInv_stale.asp"));
 <% end if %>
 <% if hasRole("AccountManager") then %>
  aux2=insFld(aux1, gFld("Other Reports", linkPrefix+"amReports/content/"));
  insDoc(aux2, gLnk(2, "SARKs by Title", linkPrefix+"amReports/content/consTitles.asp"));
  insDoc(aux2, gLnk(2, "Clients: S/R & A/M", linkPrefix+"amReports/content/client_sr_am.asp"));
  insDoc(aux2, gLnk(2, "Clients by A/M", linkPrefix+"amReports/content/client_am.asp"));
  insDoc(aux2, gLnk(2, "Intranet Survey", linkPrefix+"amReports/content/intranet_survey.asp"));
  <% if hasRole("WebMaster") then %>
   insDoc(aux2, gLnk(2, "Employee Security", linkPrefix+"amReports/content/employeeSecurity.asp")); 
  <% end if %>
 <% end if %>
<% end if %>

aux1 = insFld(mnuRoot, gFld("Directory", linkPrefix+"directory/content/", "./menu/directory.gif"));
<% if allowed(navDirectory, "sark_address_book") then %>
  insDoc(aux1, gLnk(2,"Address Book", linkPrefix+"directory/content/sark_address_book.asp"));
<% end if %>
<% if allowed(navDirectory, "employee") then %>
  insDoc(aux1, gLnk(2,"Alphabetical", linkPrefix+"directory/content/employee.asp"));
<% end if %>
<% if allowed(navDirectory, "emp_pics") then %>
  insDoc(aux1, gLnk(2,"Alpha w/ pics", linkPrefix+"directory/content/emp_pics.asp"));
<% end if %>
<% if hasRole("Sales") and allowed(navDirectory, "search") then %>
  insDoc(aux1, gLnk(2,"Search Skills", linkPrefix+"directory/content/Search_Skills.asp"));
<% end if %>
<% if allowed(navDirectory, "branches") then %>
  insDoc(aux1, gLnk(2,"Branches", linkPrefix+"directory/content/branches.asp"));
<% end if %>
<% if allowed(navDirectory, "clients") then %>
  insDoc(aux1, gLnk(2,"Client List", linkPrefix+"directory/content/clients.asp"));
<% end if %>
<% if not Session("isGuest") and allowed(navDirectory, "details") then %>
  insDoc(aux1, gLnk(2,"My Profile", linkPrefix+"directory/content/details.asp?EmpID=<%=Session("ID")%>"));
<% end if %>


<% if not Session("isGuest") then %>
aux1 = insFld(mnuRoot, gFld("Email", linkPrefix+"email/content/", "./menu/email.gif"));
<% if allowed(navEmail, "email") then %>
  insDoc(aux1, gLnk(2,"Mailbox", linkPrefix+"email/content/email.asp"));
<% end if %>

<% end if %>

aux1 = insFld(mnuRoot, gFld("Events", linkPrefix+"events/content/", "./menu/events.gif"));
<% if allowed(navEvents, "bydate") then %>
  insDoc(aux1, gLnk(2,"By Date", linkPrefix+"events/content/bydate.asp?date=<%=Server.HTMLEncode(Date)%>"));
<% end if %>
<% if allowed(navEvents, "bytype") then %>
  insDoc(aux1, gLnk(2,"By Type", linkPrefix+"events/content/bytype.asp?date=<%=Server.HTMLEncode(Date)%>"));
<% end if %>
<% if allowed(navEvents, "calendar") then %>
  insDoc(aux1, gLnk(2,"Calendar", linkPrefix+"events/content/calendar.asp?date=<%=Server.HTMLEncode(Date)%>"));
<% end if %>
<% if allowed(navEvents, "aroundtown") then %>
  insDoc(aux1, gLnk(2,"Links Around Town", linkPrefix+"events/content/aroundtown.asp"));
<% end if %>
<% if not Session("isGuest") and allowed(navEvents, "addnew") then %>
  insDoc(aux1, gLnk(2,"Add New Event", linkPrefix+"events/content/addnew.asp?date=<%=Server.HTMLEncode(Date)%>"));
<% end if %>
<% if not Session("isGuest") and allowed(navEvents, "myevents") then %>
  insDoc(aux1, gLnk(2,"My Events", linkPrefix+"events/content/myevents.asp"));
<% end if %>

aux1 = insFld(mnuRoot, gFld("News", linkPrefix+"news/content/", "./menu/news.gif"));
<% if allowed(navNews, "classifieds") then %>
  insDoc(aux1, gLnk(2,"Classifieds", linkPrefix+"news/content/classifieds.asp"));
<% end if %>
<% if allowed(navNews, "article") then %>
  insDoc(aux1, gLnk(2,"Current Article", linkPrefix+"news/content/article.asp"));
<% end if %>
<% if allowed(navNews, "article_archive") then %>
  insDoc(aux1, gLnk(2,"Article Archive", linkPrefix+"news/content/article_archive.asp"));
<% end if %>
<% if allowed(navNews, "submit_article") then %>
  insDoc(aux1, gLnk(2,"Submit Article", linkPrefix+"news/content/submit.asp"));
<% end if %>

aux1 = insFld(mnuRoot, gFld("Office", linkPrefix+"office/content/", "./menu/office.gif"));
<% if allowed(navOffice, "documentlist") then %>
  insDoc(aux1, gLnk(2,"Documents", linkPrefix+"office/content/documentlist.asp"));
<% end if %>
<% if allowed(navOffice, "album") then %>
  insDoc(aux1, gLnk(2,"Photo Album", linkPrefix+"office/content/album.asp"));
<% end if %>
<% if allowed(navOffice, "documentviewers") then %>
  insDoc(aux1, gLnk(2,"Document Viewers", linkPrefix+"office/content/documentviewers.asp"));
<% end if %>
<% if allowed(navOffice, "holiday") then %>
  insDoc(aux1, gLnk(2,"Holiday Schedule", linkPrefix+"office/content/holiday.asp"));
<% end if %>
<% if allowed(navOffice, "stats") then %>
  insDoc(aux1, gLnk(2,"Branch Statistics", linkPrefix+"office/content/stats.asp"));
<% end if %>
<% if allowed(navOffice, "corp_intranet") then %>
  insDoc(aux1, gLnk(2,"Corporate Intranet", linkPrefix+"office/content/corpintranet.asp"));
<% end if %>
<% if allowed(navOffice, "intranet_arch") then %>
  insDoc(aux1, gLnk(2,"Intranet Architecture", linkPrefix+"office/content/Intranet_Architecture.asp"));
<% end if %>

aux1 = insFld(mnuRoot, gFld("Recruiting", linkPrefix+"recruiting/content/", "./menu/recruiting.gif"));
<% if allowed(navRecruiting, "background") then %>
  insDoc(aux1, gLnk(2,"Recruiter Backgrounds", linkPrefix+"recruiting/content/Background.asp"));
<% end if %>
<% if allowed(navRecruiting, "schedule") then %>
  insDoc(aux1, gLnk(2,"Campus Schedule", linkPrefix+"recruiting/content/fall_schedule.asp"));
<% end if %>
<% if allowed(navRecruiting, "referral") then %>
  insDoc(aux1, gLnk(2,"Referral Program", linkPrefix+"recruiting/content/Referral_Program.asp"));
<% end if %>
<% if allowed(navRecruiting, "schools") then %>
  insDoc(aux1, gLnk(2,"Targeted Schools", linkPrefix+"recruiting/content/Schools.asp"));
<% end if %>


aux1 = insFld(mnuRoot, gFld("Projects", linkPrefix+"projects/content/", "./menu/projects.gif"));
<% if allowed(navProjects, "clients") then %>
  insDoc(aux1, gLnk(2,"Client summaries", linkPrefix+"projects/content/clients.asp"));
<% end if %>
<% if allowed(navProjects, "details") then %>
	<%
	set rs = DBQuery("SELECT Client.Client_ID, Client.ClientName FROM Client INNER JOIN Employee ON Client.Client_ID = Employee.Client_ID WHERE (Employee.EmployeeID = " & Session("ID") & ") AND (Client.SortOverride = 0) ORDER BY Client.ClientName")
	while not rs.eof
	    %>
		insDoc(aux1, gLnk(2,"<%=rs("ClientName")%>", linkPrefix+"projects/content/details.asp?clientid=<%=rs("Client_ID")%>"));
		<%
		rs.movenext
	wend
end if %>

aux1 = insFld(mnuRoot, gFld("Technology", linkPrefix+"technology/content/", "./menu/technology.gif"));
<% if allowed(navTechnology, "usergroups") then %>
  insDoc(aux1, gLnk(2,"User Groups", linkPrefix+"technology/content/usergroups.asp"));
<% end if %>
<% if allowed(navTechnology, "knowledge") then %>
  insDoc(aux1, gLnk(2,"Knowledge Summaries", linkPrefix+"technology/content/Knowledge_snapshot.asp"));
<% end if %>


aux1 = insFld(mnuRoot, gFld("Training", linkPrefix+"training/content/", "./menu/training.gif"));
<% if allowed(navTraining, "courseinfo") then %>
  insDoc(aux1, gLnk(2,"Training Class Info", linkPrefix+"training/content/training_archive.asp"));
<% end if %>
<% if allowed(navTraining, "downloads") then %>
  insDoc(aux1, gLnk(2,"Material Downloads", linkPrefix+"training/content/materials.asp"));
<% end if %>
<% if allowed(navTraining, "certlinks") then %>
  insDoc(aux1, gLnk(2,"Certification Links", linkPrefix+"training/content/cert_links.asp"));
<% end if %>
<% if allowed(navTraining, "certstudy") then %>
  insDoc(aux1, gLnk(2,"Cert. Notes/Tips", linkPrefix+"training/content/cert_study.asp"));
<% end if %>
<% if allowed(navTraining, "knowledge") then %>
  insDoc(aux1, gLnk(2,"Knowledge Summary", linkPrefix+"training/content/Knowledge_snapshot.asp"));
<% end if %>
<% if allowed(navTraining, "submit_training_info") then %>
  insDoc(aux1, gLnk(2,"Submit Training Info", linkPrefix+"training/content/submit.asp"));
<% end if %>



aux1 = insFld(mnuRoot, gFld("Sports", linkPrefix+"sports/content/", "./menu/sports.gif"));
<% if not Session("IsGuest") and allowed(navSports, "wellness") then %>
  insDoc(aux1, gLnk(2,"Wellness Program", linkPrefix+"wellness/content/default.asp"));
<% end if %>
<% if allowed(navSports, "racquetball") then %>
  insDoc(aux1, gLnk(2,"Racquetball", linkPrefix+"racquetball/content/about.asp"));
<% end if %>
<% if allowed(navSports, "activeSports") then 
 sql = "SELECT s.Sport_ID, Sport_Name " & _
       "FROM Sports s, Sports_Season ss " & _
       "WHERE s.sport_id = ss.sport_id AND Season_Start <= (GETDATE()+14) " & _
       "AND Season_End >= (GETDATE()-14) " & _
       "ORDER BY Sport_Name"
 set rsSports = DBQuery(sql)
 do while not rsSports.eof
  %>
   insDoc(aux1, gLnk(2,"<%=rsSports("Sport_Name")%>", linkPrefix+"sports/content/showSport.asp?sportID=<%=rsSports("Sport_ID")%>"));
  <%
  rsSports.moveNext
 loop
end if %>


<% if hasRole("SolutionServices") or hasRole("WebMaster") then %>
aux1 = insFld(mnuRoot, gFld("Repository", linkPrefix+"SolutionServices/content", "./menu/solutionServices.gif"));
	insDoc(aux1, gLnk(2,"About", linkPrefix+"SolutionServices/content/"));
	insDoc(aux1, gLnk(2,"Doc. Repository", linkPrefix+"SolutionServices/content/view_repository.asp"));
<% end if %>
