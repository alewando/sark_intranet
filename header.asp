<%
'-------------------------
'  Get current page info  
'-------------------------
dim curpage, i, startTime, curTitle
curTitle = sectionTitle
if pageTitle <> "" then curTitle = curTitle & " - " & pageTitle
startTime = now
curpage = DefaultNavPage
i = instr(curpage, "=")
while i > 0
	curpage = left(curpage, i-1) & "~" & mid(curpage, i+1)
	i = instr(curpage, "=")
	wend
%>


<html>

<head>
<title><%=curTitle%></title>

<script language=javascript>
<!--
var pageLoaded = false
var browserName = navigator.appName
var browserAbbr = "";
var browserVersion = parseInt(navigator.appVersion.substring(0,1))
var curPage = "<%=curPage%>"
var curSection = "'<%=sectionTitle%>'"
if (browserName=="Microsoft Internet Explorer") {browserAbbr = "IE";}
else if (browserName=="Netscape") {browserAbbr = "NS";}

function statmsg(msg){
	top.status = msg
	}

function selectNavItem(){
	var sel = document.frmNav.selNav
	window.location.href = sel.options[sel.selectedIndex].value
	}

function setCookie(key, val){
	var today = new Date()
	var expires = new Date()
	expires.setTime(today.getTime() + 1000*60*60*24*365)
	document.cookie = "<%=SectionDir%>" + key + "=" + val + ";path=/<%=Application("web")%>;expires=" + expires.toGMTString()
	}

function SlideShow(title, dir){
	winSlides = window.open('/<%=Application("web")%>/extras/slideshow.asp?title=' + title + '&dir=' + dir,'Slideshow','height=400,width=500,toolbar=0,location=0,directories=0,status=0,menubar=0,resizable=0')
	}
// -->
</script>

<!-- #include file="style.htm" -->

</head>


<basefont color="black" face="ms sans serif, arial, geneva">

<body background="../../common/images/<%=sectionImg%>" onLoad="pageLoaded = true">


<table cellpadding=0 cellspacing=0><tr>

<!-- BODY -->
<td width=490 valign=top align=left>


<!-- HEADER -->
<table border=0 width=490 cellpadding=0 cellspacing=0>
	<form name="frmNav">
	<tr>
		<td nowrap><font size=5 face="times"><%=sectionTitle%></font></td>
		<td align=right valign=bottom>
<%
'--------------------------------
'  Display drop down navigation  
'--------------------------------
if sectionNav <> "" then
	if instr(SectionNav, "<option") > 0 then
		response.write("		<select name=selNav onChange='selectNavItem()'>" & NL)
		response.write(SectionNav)
		if instr(SectionNav, "selected") = 0 then
			'SectionBtns = ""	'Disable "set default" button if current page not in nav list
			response.write("			<option selected>" & pageTitle & "</option>" & NL)
			end if
		response.write("		</select>" & NL)
	else
		response.write(SectionNav)
		end if
	end if
%>
		</td>
	</tr>
	<tr><td colspan=2><hr size=1></td></tr>
	</form>

	<tr><td colspan=2>

<table border=0 cellpadding=5 cellspacing=0 width="100%"><tr><td><font size=1 face="ms sans serif, arial, geneva">










<!-- START OF BODY -->

