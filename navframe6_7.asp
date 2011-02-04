<!-- #include file="script.asp" -->

<%
'---------------------------------------------------------------------------------
' Description:	Displays user-specific navigation hierarchy.
' History:		06/02/1999 - KDILL - Created
'				08/20/1999 - KDILL - Modified for preferences change.
'				01/12/2000 - KDILL - Added tech alt page check.
'---------------------------------------------------------------------------------



'-----------------------------------
'	Get navigational preferences	
'-----------------------------------
dim navWelcome, navDirectory, navEmail, navEvents, navNews, navOffice, navProjects, navTechnology

function getNavPref(cookieKey, defaultVal)
	dim result
	result = request.cookies(cookieKey)
	if result = "" then result = defaultVal
	getNavPref = "*" & result & "*"
end function

function allowed(srcVal, searchVal)
	allowed = (InStr(srcVal, "*" & searchVal & "*") > 0)
end function

navWelcome = getNavPref("navWelcome", Application("DefaultNavWelcome"))
navDirectory = getNavPref("navDirectory", Application("DefaultNavDirectory"))
navEmail = getNavPref("navEmail", Application("DefaultNavEmail"))
navEvents = getNavPref("navEvents", Application("DefaultNavEvents"))
navNews = getNavPref("navNews", Application("DefaultNavNews"))
navOffice = getNavPref("navOffice", Application("DefaultNavOffice"))
navProjects = getNavPref("navProjects", Application("DefaultNavProjects"))
navTechnology = getNavPref("navTechnology", Application("DefaultNavTechnology"))
navTraining = getNavPref("navTraining", Application("DefaultNavTraining"))
navSports = getNavPref("navSports", Application("DefaultNavSports"))
%>

<html>

<head>
<title>Main Navigation</title>

<!-- #include file="style.htm" -->


<script language="javascript">
<!--
function showWeather(){
	var winWeather
	winWeather = window.open("/IntranetDialogs/welcome/weatherMaps.asp?index=0","weatherMap","height=330,width=470,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=0")
	return false;
	}

function sendEmail(){
	var winTech
	winTech = window.open('/IntranetDialogs/email.asp?editto=yes', 'SendEmail','height=330,width=520,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=1, scrollbar=1');
	return false;
	}

// -->
</script>

</head>


<base target="body">


<body background="common/images/starback.jpg" bgcolor=black text=white link=white alink=white vlink=white><center>

<a href="/" target="_top"><img src="common/images/nav/logo2.gif" width=135 height=99 alt="Sark Cincinnati" border=0></a><br>


<table border=0 cellpadding=0 cellspacing=0><tr><td><font face="arial, geneva, helvetica" size=2 color=white>


<!-- Welcome -->
<a class=menu href="welcome/content/"><b>Welcome</b></a><font size=1>
<% if allowed(navWelcome, "comics") then %>
	<br>&nbsp;
		<a class=submenu href="welcome/content/humor.asp">
		Humor</a>
<% end if %>
<% if allowed(navWelcome, "preferences") then %>
	<br>&nbsp;
		<a class=submenu href="preferences/content/default.asp">
		Preferences</a>
<% end if %>
	</font>


<!-- Directory -->
<br><a class=menu href="directory/content/"><b>Directory</b></a><font size=1>
<% if allowed(navDirectory, "sark_address_book") then %>
	<br>&nbsp;
		<a class=submenu href="directory/content/sark_address_book.asp">
		Address Book</a>
<% end if %>
<% if allowed(navDirectory, "employee") then %>
	<br>&nbsp;
		<a class=submenu href="directory/content/employee.asp">
		Alphabetical</a>
<% end if %>
<% if allowed(navDirectory, "emp_pics") then %>
	<br>&nbsp;
		<a class=submenu href="directory/content/emp_pics.asp">
		Alphabetical w/ pics</a>
<% end if %>
<% if allowed(navDirectory, "branches") then %>
	<br>&nbsp;
		<a class=submenu href="directory/content/branches.asp">
		Branches</a>
<% end if %>
<% if allowed(navDirectory, "clients") then %>
	<br>&nbsp;
		<a class=submenu href="directory/content/clients.asp">
		Client list</a>
<% end if %>
<%if Session("ID") <> "" and allowed(navDirectory, "details") then%>
	<br>&nbsp;
		<a class=submenu href="directory/content/details.asp?EmpID=<% =Session("ID") %>">
		My profile</a>
<%end if%>
	</font>


<!-- Email -->
<br><a class=menu href="email/content/"><b>Email</b></a><font size=1>
<% if allowed(navEmail, "email") then %>
	<br>&nbsp;
		<a class=submenu href="email/content/email.asp">
		Mailbox</a>
<% end if %>
	</font>


<!-- Events -->
<br><a class=menu href="events/content/"><b>Events</b></a><font size=1>
<% if allowed(navEvents, "bydate") then %>
	<br>&nbsp;
		<a class=submenu href="events/content/bydate.asp?date=<% =Server.HTMLEncode(Date) %>">
		By date</a>
<% end if %>
<% if allowed(navEvents, "bytype") then %>
	<br>&nbsp;
		<a class=submenu href="events/content/bytype.asp?date=<% =Server.HTMLEncode(Date) %>">
		By type</a>
<% end if %>
<% if allowed(navEvents, "calendar") then %>
	<br>&nbsp;
		<a class=submenu href="events/content/calendar.asp?date=<% =Server.HTMLEncode(Date) %>">
		Calendar</a>
<% end if %>
<% if allowed(navEvents, "addnew") then %>
	<br>&nbsp;
		<a class=submenu href="events/content/addnew.asp?date=<% =Server.HTMLEncode(Date) %>">
		Add new event</a>
<% end if %>
<% if allowed(navEvents, "myevents") then %>
	<br>&nbsp;
		<a class=submenu href="events/content/myevents.asp">
		My Events</a>
<% end if %>
	</font>

	
<!-- News -->
<br><a class=menu href="news/content/"><b>News</b></a><font size=1>
<% if allowed(navNews, "classifieds") then %>
	<br>&nbsp;
		<a class=submenu href="news/content/classifieds.asp">
		Classifieds</a>
<% end if %>
<% if allowed(navNews, "article") then %>
	<br>&nbsp;
		<a class=submenu href="news/content/article.asp">
		Current article</a>
<% end if %>
<% if allowed(navNews, "article_archive") then %>
	<br>&nbsp;
		<a class=submenu href="news/content/article_archive.asp">
		Article archive</a>
<% end if %>
	</font>


<!-- Office -->
<br><a class=menu href="office/content/"><b>Office</b></a><font size=1>
<% if allowed(navOffice, "documentlist") then %>
	<br>&nbsp;
		<a class=submenu href="office/content/documentlist.asp">
		Documents</a>
<% end if %>
<% if allowed(navOffice, "holiday") then %>
	<br>&nbsp;
		<a class=submenu href="office/content/holiday.asp">
		Holiday schedule</a>
<% end if %>
<% if allowed(navOffice, "stats") then %>
	<br>&nbsp;
		<a class=submenu href="office/content/stats.asp">
		Statistics</a>
<% end if %>
	</font>

<!-- Projects -->	
<br><a class=menu href="projects/content/"><b>Projects</b></a><font size=1>
<% if allowed(navProjects, "clients") then %>
	<br>&nbsp;&nbsp;
		<a class=submenu href="projects/content/clients.asp">
		Client summaries</a>
<% end if
if (Session("ID")) and allowed(navProjects, "details") then
	set rs = DBQuery("SELECT Client.Client_ID, Client.ClientName FROM Client INNER JOIN Employee ON Client.Client_ID = Employee.Client_ID WHERE (Employee.EmployeeID = " & Session("ID") & ") AND (Client.SortOverride = 0) ORDER BY Client.ClientName")
	while not rs.eof
%>
	<br>&nbsp;
		<a class=submenu href="projects/content/details.asp?clientid=<% =rs("Client_ID") %>&">
		&nbsp;<%=rs("ClientName") %></a>
<%
		rs.movenext
		wend
end if
%>
	</font>

<!-- Technology -->
<br><a class=menu href="technology/content/"><b>Technology</b></a><font size=1>

<% if allowed(navTechnology, "usergroups") then %>
	<br>&nbsp;&nbsp;
		<a class=submenu href="technology/content/usergroups.asp">
		User Groups</a>
<% end if %>

</font>

<!-- Training -->
<br><a class=menu href="training/content/about_training.asp"><b>Training &<br>Certifications</b></a><font size=1>

<% if allowed(navTraining, "courseinfo") then %>
	<br>&nbsp;&nbsp;
		<a class=submenu href="training/content/training_archive.asp">
		Training Class Info</a>
<% end if %>
<% if allowed(navTraining, "downloads") then %>
	<br>&nbsp;&nbsp;
		<a class=submenu href="training/content/materials.asp">
		Material Downloads</a>
<% end if %>
<% if allowed(navTraining, "cbtcourses") then %>
	<br>&nbsp;&nbsp;
		<a class=submenu href="training/content/training_cbt.asp">
		CBT Courses</a>
<% end if %>
<% if allowed(navTraining, "certlinks") then %>
	<br>&nbsp;&nbsp;
		<a class=submenu href="training/content/cert_links.asp">
		Certifications, Links</a>
<% end if %>
<% if allowed(navTraining, "certstudy") then %>
	<br>&nbsp;&nbsp;
		<a class=submenu href="training/content/cert_study.asp">
		Certifications, Notes/Tips</a>
<% end if %>
<% if allowed(navTraining, "knowledge") then %>
	<br>&nbsp;&nbsp;
		<a class=submenu href="training/content/Knowledge_snapshot.asp">
		Knowledge Summary</a>
<% end if %>



</font>

<!-- Sports -->
<br><a class=menu href="sports/content/"><b>Sports</b></a><font size=1>

<% if allowed(navSports, "wellness") then %>
	<br>&nbsp;&nbsp;
		<a class=submenu href="http://splash.sarkcincinnati.com/wellness" target="blank">
		Wellness Program</a>
<% end if %>

<% if allowed(navSports, "basketball") then %>
	<br>&nbsp;&nbsp;
		<a class=submenu href="sports/content/basketball.asp">
		Basketball League</a>
<% end if %>

<% if allowed(navSports, "softball") then %>
	<br>&nbsp;&nbsp;
		<a class=submenu href="sports/content/softball.asp">
		Softball League</a>
<% end if %>

<% if allowed(navSports, "golf") then %>
	<br>&nbsp;&nbsp;
		<a class=submenu href="sports/content/Golf_League.asp">
		Golf League</a>
<% end if %>

	</font>


<!-- Tools -->
<% if Session("Username") = "cdolan" or _
	session("Username") = "dmarston" then %>
<br><a class=menu href="tools/content/"><b>Tools</b></a><font size=1>
<br>&nbsp;&nbsp;<a class=submenu href="tools/content/add_emp.asp">Add Employee</a>
<br>&nbsp;&nbsp;<a class=submenu href="tools/content/chg_emp.asp">Edit Employee</a>
<br>&nbsp;&nbsp;<a class=submenu href="tools/content/del_emp.asp">Delete Employee</a>
<br>&nbsp;&nbsp;<a class=submenu href="tools/content/add_client.asp">Add Client</a>
<br>&nbsp;&nbsp;<a class=submenu href="tools/content/edit_client.asp">Edit Client</a>
<br>&nbsp;&nbsp;<a class=submenu href="tools/content/approve_postings.asp">Approve Postings</a>
<br>&nbsp;&nbsp;<a class=submenu href="tools/content/techapproval.asp">Approve New Techs</a>
	</font>
<% end if %>

</font></td></tr></table>

</center></body>

</html>
