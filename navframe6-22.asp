<!-- #include file="script.asp" -->

<html>

<head>
<title>Main Navigation</title>

<!-- #include file="style.htm" -->


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

<%
		Response.Write("<SCRIPT SRC=""menu/menus.js""></SCRIPT>")
		Response.Write("<LINK REL=""stylesheet"" TYPE=""text/css"" HREF=""menu/menus.css""></SCRIPT>")
%>
<!-- #include file='menu/menus.asp' -->

</head>


<base target="body">


<body background="common/images/starback.jpg" bgcolor=black text=white link=white alink=white vlink=white leftmargin=0><center>

<a href="http://www.sark.com" target="_top"><img src="common/images/nav/logo2.gif" width=135 height=99 alt="New Sark.com Site" border=0></a><br>

<!-- Menu code by Ryan Dlugsz 6/2/00
		This code implements a DHTML menu for the left hand side
		Code is in 'menu' subdirectory
-->
<!-- #include file="menu/menu.asp" -->

<SCRIPT>
<!--
top.body.location="welcome/content/default.asp";
// -->
</SCRIPT>

</body>

</html>
