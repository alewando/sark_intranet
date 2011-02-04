
<!-- #include file="./script.asp" -->
<html>

<head>
<title>Main Navigation</title>

<!-- #include file="./style.htm" -->

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
navRecruiting = getNavPref("navRecruiting", Application("DefaultNavRecruiting"))
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

<link rel="stylesheet" TYPE="text/css" HREF="./menu/ftie4style.css">
<script src="./menu/ftiens4.js"></script>

<script language="JavaScript">
<!-- #include file="./menu/menu.asp" -->
</script>
</head>


<base target="body">


<body background="./common/images/logo_bg.jpg" bgcolor=black text=white link=white alink=white vlink=white leftmargin=0>
<a href="http://www.sark.com" target="_top"><img name="imgLogo" src="./common/images/nav/sarkcincinnati.gif" width=150 height=72 alt="New Sark.com Site" border=0></a><br>
<script>
 initializeDocument();
</script>

</body>

</html>
