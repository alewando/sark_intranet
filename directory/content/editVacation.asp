<!--
Developers:   SSeissiger, KWahoff
Date:         07/19/2000
Description:  Displays listing of employee vacation information
-->


<!-- #include file="../../script.asp" -->

<html>
<head>
<title>Add employee vacation info</title>
</head>

<script language=javascript>
<!--
var browserAbbr = "";
var browserName = navigator.appName
var browserVersion = parseInt(navigator.appVersion.substring(0,1))
if (browserName=="Microsoft Internet Explorer") {browserAbbr = "IE";}
else if (browserName=="Netscape") {browserAbbr = "NS";}

function closeWindow(){
	//if (self.opener && ((browserAbbr=="IE" && browserVersion > 3) || (browserAbbr=="NS" && browserVersion > 2))) self.opener.location.reload();
	window.close();
	}
-->
</script>

<!-- 'closeWindow()' -->
<body bgcolor=silver onLoad=window.close()>

<!-- 
<font size=1 face="ms sans serif, arial, geneva">Saving information...</font>
-->

<%
ls_empid = Request("EmployeeID")

	ls_vacstart = 	Clean(trim(Request("txtVacation_start")))
	ls_vacend = 	Clean(trim(Request("txtVacation_end")))
	ls_vaccomments = Clean(trim(Request("txtComments")))
	
	set rs = DBQuery("INSERT into vacation (vacation_start, vacation_end, comments, employee_id) Values('" & ls_vacstart & "', '" & ls_vacend & "', '" & ls_vaccomments & "', '" & ls_empid & "' )")
         
%>

</body>
</html>
-->