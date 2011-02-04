<!--
Developer:    Sark Beach
Changed:      JLove
Date:         10/18/2000
Description:  User preference page.
-->
<%
'Response.CacheControl = "no-cache"
'Response.AddHeader "Pragma", "no-cache"
'Response.Expires = -1
%>
<!-- #include file="../section.asp" -->


<script language=javascript>
<!--
function setPref(key, val){
	var today = new Date()
	var expires = new Date()
	expires.setTime(today.getTime() + 1000*60*60*24*365)
	document.cookie = key + "=" + val + ";path=/intranet;expires=" + expires.toGMTString()
	}

function savePrefs(){
	
	var cookieWelcome = ""
	var cookieDirectory = ""
	var cookieEmail = ""
	var cookieEvents = ""
	var cookieNews = ""
	var cookieOffice = ""
	var cookieRecruiting = ""
	var cookieProjects = ""
	var cookieTechnology = ""
	var cookieTraining = ""
	var cookieSports = ""

	var tempWelcome = ""
	var tempDirectory = ""
	var tempEmail = ""
	var tempEvents = ""
	var tempNews = ""
	var tempOffice = ""
	var tempRecruiting = ""
	var tempProjects = ""
	var tempTechnology = ""
	var tempTraining = ""
	var tempSports = ""

	for (i=0; i < document.frmInfo.elements.length; i++){
		ctrl = document.frmInfo.elements[i]
		ctrlnm = ctrl.name
		if (ctrlnm.substring(0,3) == "nav"){
			j = ctrlnm.indexOf("_")
			ctrlkey = ctrlnm.substring(3,j)
			ctrlval = ctrlnm.substring(j+1)
			if (((ctrl.checked) && (ctrl.value=="+")) || ((!ctrl.checked) && (ctrl.value=="-"))){
				if (ctrlkey == "Welcome") cookieWelcome += "*" + ctrlval
				if (ctrlkey == "Directory") cookieDirectory += "*" + ctrlval
				if (ctrlkey == "Email") cookieEmail += "*" + ctrlval
				if (ctrlkey == "Events") cookieEvents += "*" + ctrlval
				if (ctrlkey == "News") cookieNews += "*" + ctrlval
				if (ctrlkey == "Office") cookieOffice += "*" + ctrlval
				if (ctrlkey == "Recruiting") cookieRecruiting += "*" + ctrlval
				if (ctrlkey == "Projects") cookieProjects += "*" + ctrlval
				if (ctrlkey == "Technology") cookieTechnology += "*" + ctrlval
				if (ctrlkey == "Training") cookieTraining += "*" + ctrlval
				if (ctrlkey == "Sports") cookieSports +="*" + ctrlval
				}
			}
		}

	cookieWelcome += "*"
	cookieDirectory += "*"
	cookieEmail += "*"
	cookieEvents += "*"
	cookieNews += "*"
	cookieOffice += "*"
	cookieRecruiting += "*"
	cookieProjects += "*"
	cookieTechnology += "*"
	cookieTraining += "*"
	cookieSports += "*"

//	alert("cookieWelcome = '" + cookieWelcome + "'")
//	alert("cookieDirectory = '" + cookieDirectory + "'")
//	alert("cookieEmail = '" + cookieEmail + "'")
//	alert("cookieEvents = '" + cookieEvents + "'")
//	alert("cookieNews = '" + cookieNews + "'")
//	alert("cookieOffice = '" + cookieOffice + "'")
//	alert("cookieProjects = '" + cookieProjects + "'")
//	alert("cookieTechnology = '" + cookieTechnology + "'")
//  alert("cookieTraining = '" + cookieTraining + "'")
//  alert("cookieSports = '" + cookieSports + "'")

	setPref("navWelcome", cookieWelcome)
	setPref("navDirectory", cookieDirectory)
	setPref("navEmail", cookieEmail)
	setPref("navEvents", cookieEvents)
	setPref("navNews", cookieNews)
	setPref("navOffice", cookieOffice)
	setPref("navRecruiting", cookieRecruiting)
	setPref("navProjects", cookieProjects)
	setPref("navTechnology", cookieTechnology)
	setPref("navTraining", cookieTraining)
	setPref("navSports", cookieSports)

	window.parent.navmain.location.reload()
	alert("Preferences saved.")
	}
	
function ClearPrefs(){

	for (i=0; i < document.frmInfo.elements.length; i++){
			var ctrl = document.frmInfo.elements[i]
			if (ctrl.checked){
				ctrl.checked = false
			}
		}
	}
	
// -->
</script>


<%
function val(srcVal, searchVal)
	val = (InStr(srcVal, "*" & searchVal & "*") > 0)
end function

function checked(srcVal, searchVal)
	checked = ""
	if val(srcVal, searchVal) then checked = "checked"
end function

function notchecked(srcVal, searchVal)
	notchecked = "checked"
	if val(srcVal, searchVal) then notchecked = ""
end function

'-----------------------------------
'	Try to get values from cookies	
'-----------------------------------
navWelcome		= request.cookies("navWelcome")
navDirectory	= request.cookies("navDirectory")
navEmail		= request.cookies("navEmail")
navEvents		= request.cookies("navEvents")
navNews			= request.cookies("navNews")
navOffice		= request.cookies("navOffice")
navRecruiting	= Request.Cookies("navRecruiting")
navProjects		= request.cookies("navProjects")
navTechnology	= request.cookies("navTechnology")
navTraining		= Request.Cookies("navTraining")
navSports		= request.cookies("navSports")


'-------------------------------------------------------
'	If no values stored in cookies, use default values	
'-------------------------------------------------------
if navWelcome = "" then navWelcome = Application("DefaultNavWelcome")
if navDirectory	= "" then navDirectory = Application("DefaultNavDirectory")
if navEmail = "" then navEmail = Application("DefaultNavEmail")
if navEvents = "" then navEvents = Application("DefaultNavEvents")
if navNews = "" then navNews = Application("DefaultNavNews")
if navOffice = "" then navOffice = Application("DefaultNavOffice")
if navRecruiting ="" then navRecruiting = Application("DefaultNavRecruiting")
if navProjects = "" then navProjects = Application("DefaultNavProjects")
if navTechnology = "" then navTechnology = Application("DefaultNavTechnology")
if navTraining = "" then navTraining = Application("DefaultNavTraining")
if navSports = "" then navSports = Application("DefaultNavSports")


'response.write("navWelcome = '" & navWelcome & "'<br>")
'response.write("navDirectory = '" & navDirectory & "'<br>")
'response.write("navEmail = '" & navEmail & "'<br>")
'response.write("navEvents = '" & navEvents & "'<br>")
'response.write("navNews = '" & navNews & "'<br>")
'response.write("navOffice = '" & navOffice & "'<br>")
'response.write("navProjects = '" & navProjects & "'<br>")
'response.write("navTechnology = '" & navTechnology & "'<br>")
'response.write("navTraining = '" & navTraining & "'<br>")
'response.write("navSports = '" & navSports & "'<br>")
%>



<center><form name="frmInfo">


<table border=0 width="90%"><tr><td><font size=1 face="ms sans serif, arial, geneva">
Here are your preferences...
Make your changes and click "Save Preferences" to save them to a cookie on your
computer and reload the navigation menu.
</font></td></tr></table><br>


<table class=tableShadow border=0 cellpadding=10 cellspacing=0><tr>

<td valign=top bgcolor=#ffffcc nowrap><font size=1 face="ms sans serif, arial, geneva">

	<p><b>Welcome</b>
<!--	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navWelcome_comics" value="+" >Humor-->
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navWelcome_preferences" value="+" <%=checked(navWelcome, "preferences")%>>Preferences

	<p><b>Directory</b>
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navDirectory_employee" value="+" <%=checked(navDirectory, "employee")%>>Alphabetical
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navDirectory_emp_pics" value="+" <%=checked(navDirectory, "emp_pics")%>>Alphabetical w/ pics
	<% if hasRole("Sales") then %>
	 <br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navDirectory_search" value="+" <%=checked(navDirectory, "search")%>>Search Skills
	<% end if %>
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navDirectory_branches" value="+" <%=checked(navDirectory, "branches")%>>Branches
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navDirectory_clients" value="+" <%=checked(navDirectory, "clients")%>>Client list
	<% if not Session("isGuest") then %>
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navDirectory_details" value="+" <%=checked(navDirectory, "details")%>>My profile
	<% end if %>

	<% if not Session("isGuest") then %>
	<p><b>Email</b>
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navEmail_email" value="+" <%=checked(navEmail, "email")%>>MailBox
    <% end if %>
    
	<p><b>Events</b>
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navEvents_bydate" value="+" <%=checked(navEvents, "bydate")%>>By Date
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navEvents_bytype" value="+" <%=checked(navEvents, "bytype")%>>By Type
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navEvents_calendar" value="+" <%=checked(navEvents, "calendar")%>>Calendar
		<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navEvents_aroundtown" value="+" <%=checked(navEvents, "aroundtown")%>>Links Around Town
	<% if not Session("isGuest") then %>
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navEvents_addnew" value="+" <%=checked(navEvents, "addnew")%>>Add New
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navEvents_myevents" value="+" <%=checked(navEvents, "myevents")%>>My Events
	<% end if %>

	<p><b>News</b>
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navNews_classifieds" value="+" <%=checked(navNews, "classifieds")%>>Classifieds
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navNews_article" value="+" <%=checked(navNews, "article")%>>Current article
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navNews_article_archive" value="+" <%=checked(navNews, "article_archive")%>>Article archive
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navNews_submit_article" value="+" <%=checked(navNews, "submit_article")%>>Submit article
	

</font></td><td width=10 bgcolor=#ffffcc></td><td valign=top bgcolor=#ffffcc nowrap><font size=1 face="ms sans serif, arial, geneva">

	<p><b>Office</b>
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navOffice_documentlist" value="+" <%=checked(navOffice, "documentlist")%>>Documents
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navOffice_album" value="+" <%=checked(navOffice, "album")%>>SARK Photo Album
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navOffice_documentviewers" value="+" <%=checked(navOffice, "documentviewers")%>>Document Viewers	
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navOffice_holiday" value="+" <%=checked(navOffice, "holiday")%>>Holiday schedule
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navOffice_stats" value="+" <%=checked(navOffice, "stats")%>>Statistics
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navOffice_corp_intranet" value="+" <%=checked(navOffice, "corp_intranet")%>>Corporate Intranet
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navOffice_intranet_arch" value="+" <%=checked(navOffice, "intranet_arch")%>>Intranet Architecture

	<p><b>Recruiting</b>
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navRecruiting_background" value="+" <%=checked(navRecruiting, "background")%>>Recruiter Backgrounds
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navRecruiting_schedule" value="+" <%=checked(navRecruiting, "schedule")%>>Campus Schedule
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navRecruiting_referral" value="+" <%=checked(navRecruiting, "referral")%>>Referral Program
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navRecruiting_schools" value="+" <%=checked(navRecruiting, "schools")%>>Targeted Schools

	<p><b>Projects</b>
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navProjects_clients" value="+" <%=checked(navProjects, "clients")%>>Client summaries
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navProjects_details" value="+" <%=checked(navProjects, "details")%>>Current client

	<p><b>Technology</b>
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navTechnology_usergroups" value="+" <%=checked(navTechnology, "usergroups")%>>User Groups
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navTechnology_knowledge" value="+" <%=checked(navTechnology, "knowledge")%>>Knowledge Summary
	
	<p><b>Training</b>
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navTraining_courseinfo" value="+" <%=checked(navTraining, "courseinfo")%>>Training Class Information
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navTraining_downloads" value="+" <%=checked(navTraining, "downloads")%>>Course Material Downloads
		<!--
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navTraining_cbtcourses" value="+" <%=checked(navTraining, "cbtcourses")%>>CBT Courses
	-->
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navTraining_certlinks" value="+" <%=checked(navTraining, "certlinks")%>>Certifications, Links
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navTraining_certstudy" value="+" <%=checked(navTraining, "certstudy")%>>Certifications, Notes/Tips
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navTraining_submit_training_info" value="+" <%=checked(navTraining, "submit_training_info")%>>Submit training info


	<p><b>Sports</b>
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navSports_activeSports" value="+" <%=checked(navSports, "activeSports")%>>Active Sports	
	<% if not Session("isGuest") then %>
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navSports_wellness" value="+" <%=checked(navSports, "wellness")%>>Wellness Program
	<% end if %>
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navSports_racquetball" value="+" <%=checked(navSports, "racquetball")%>>Racquetball
	<!--
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navSports_basketball" value="+" <%=checked(navSports, "basketball")%>>Basketball League
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navSports_softball" value="+" <%=checked(navSports, "softball")%>>Softball League
	<br>&nbsp;&nbsp;&nbsp;&nbsp;<input type=checkbox name="navSports_golf" value="+" <%=checked(navSports, "golf")%>>Golf League
    -->
</font></td></tr></table><p>

<input type=button class=button name=btnSave value="Save Preferences" onClick="savePrefs()">
<input type=button class=button name=btnReset value="Clear Preferences" onClick="ClearPrefs()">
<input type=button class=button name=btnCancel value="Cancel" onClick="window.location='../../welcome/content/default.asp'">


</form>

</center>

<!-- #include file="../../footer.asp" -->
