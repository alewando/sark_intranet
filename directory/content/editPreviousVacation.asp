<!--
Developers:   SSeissiger, KWahoff
Date:         07/19/2000
Description:  Displays listing of employee vacation information
-->


<!-- #include file="../../script.asp" -->

<script language=javascript>

var browserAbbr = "";
var browserName = navigator.appName
var browserVersion = parseInt(navigator.appVersion.substring(0,1))
if (browserName=="Microsoft Internet Explorer") {browserAbbr = "IE";}
else if (browserName=="Netscape") {browserAbbr = "NS";}

function closeWindow(){
	//if (self.opener && ((browserAbbr=="IE" && browserVersion > 3) || (browserAbbr=="NS" && browserVersion > 2))) self.opener.location.reload();
	window.close();
	}


</script>
	
<html>
<head>
<title>Update employee vacation info</title>
</head>
 
<body bgcolor=silver onLoad='window.close();'>
<%
'For testing
session("GotAcctMgr")= true 
%>
<!-- <font size=1 face="ms sans serif, arial, geneva">Saving information...</font> -->


<%
ls_empid = Request.QueryString("empid")

ls_vacid = Clean(trim(Request("txtVacation_id")))
ls_vacstart = 	Clean(trim(Request("txtVacation_start")))
ls_vacend = 	Clean(trim(Request("txtVacation_end")))
ls_vaccomments = Clean(trim(Request("txtComments")))
	
	
set rs = DBQuery("UPDATE vacation SET vacation_start = '" & ls_vacstart & "', vacation_end = '" & ls_vacend & "', comments = '" & ls_vaccomments & "' WHERE vacation_id = " & ls_vacid)
%>

</body>
</html>